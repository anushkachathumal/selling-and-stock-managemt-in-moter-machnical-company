Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports System.IO.StreamReader
'Imports Microsoft.Office.Interop.Excel
'Imports System.DateTime

Public Class frmUpdate_Status

    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim c_dataCustomer1 As System.Data.DataTable
    Dim _Customer As String
    Dim _Department As String
    Dim _Merchant As String
  
    Function Load_PO()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select M10Dis as [Department] from M10OTD_Department"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                With cboPO
                    .DataSource = M01
                    .Rows.Band.Columns(0).Width = 175
                End With
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

    Function Load_Quality()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select left(Material_Dis,5) as [Quality No] from ZPP_DEL where left(Material_Dis,1)<>'Y' group by left(Material_Dis,5) union select left(Material_Dis,8) from ZPP_DEL where left(Material_Dis,1)='Y' group by left(Material_Dis,8)"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                With cboQuality
                    .DataSource = M01
                    .Rows.Band.Columns(0).Width = 175
                End With
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

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    

    

    Function Load_Gride()
        Dim CustomerDataClass As New DAL_InterLocation
        c_dataCustomer1 = CustomerDataClass.MakeDataTableCheck_Customer
        UltraGrid2.DataSource = c_dataCustomer1
        With UltraGrid2
            .DisplayLayout.Bands(0).Columns(0).Width = 40
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 230
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            '.DisplayLayout.Bands(0).Columns(2).Width = 90
            '.DisplayLayout.Bands(0).Columns(3).Width = 90
            '  .DisplayLayout.Bands(0).Columns(4).Width = 90
            '  .DisplayLayout.Bands(0).Columns(5).Width = 90
            '  .DisplayLayout.Bands(0).Columns(6).Width = 90
            ' .DisplayLayout.Bands(0).Columns(7).Width = 80

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function

    Function Load_GrideDep()
        Dim CustomerDataClass As New DAL_InterLocation
        c_dataCustomer1 = CustomerDataClass.MakeDataTableCheck_Dep
        UltraGrid1.DataSource = c_dataCustomer1
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 40
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 230
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            '.DisplayLayout.Bands(0).Columns(2).Width = 90
            '.DisplayLayout.Bands(0).Columns(3).Width = 90
            '  .DisplayLayout.Bands(0).Columns(4).Width = 90
            '  .DisplayLayout.Bands(0).Columns(5).Width = 90
            '  .DisplayLayout.Bands(0).Columns(6).Width = 90
            ' .DisplayLayout.Bands(0).Columns(7).Width = 80

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function

    Function Load_GrideMerch()
        Dim CustomerDataClass As New DAL_InterLocation
        c_dataCustomer1 = CustomerDataClass.MakeDataTableCheck_Merch
        UltraGrid3.DataSource = c_dataCustomer1
        With UltraGrid3
            .DisplayLayout.Bands(0).Columns(0).Width = 40
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 230
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            '.DisplayLayout.Bands(0).Columns(2).Width = 90
            '.DisplayLayout.Bands(0).Columns(3).Width = 90
            '  .DisplayLayout.Bands(0).Columns(4).Width = 90
            '  .DisplayLayout.Bands(0).Columns(5).Width = 90
            '  .DisplayLayout.Bands(0).Columns(6).Width = 90
            ' .DisplayLayout.Bands(0).Columns(7).Width = 80

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function

    Function Load_Customer()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Dim I As Integer

        Try
            If Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                Sql = "select Customer from OTD_SMS where Department in ('" & _Department & "') and Merchant in ('" & _Merchant & "')   group by Customer"
            ElseIf Trim(txtDepartment.Text) <> "" Then
                Sql = "select Customer from OTD_SMS where Department in ('" & _Department & "') group by Customer"
            ElseIf Trim(txtMerchant.Text) <> "" Then
                Sql = "select Customer from OTD_SMS where Merchant in ('" & _Merchant & "') group by Customer"
            Else
                Sql = "select Customer from OTD_SMS group by Customer"
            End If

            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            I = 0

            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                'For Each DTRow1 As DataRow In M01.Tables(0).Rows
                newRow("##") = False
                newRow("Customer Name") = M01.Tables(0).Rows(I)("Customer")

                c_dataCustomer1.Rows.Add(newRow)
                I = I + 1

            Next
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

    Function Load_Department()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Dim I As Integer

        Try
            If Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                Sql = "select Department from OTD_SMS where Customer in ('" & _Customer & "') and Merchant in ('" & _Merchant & "')   group by Department"
            ElseIf Trim(txtCustomer.Text) <> "" Then
                Sql = "select Department from OTD_SMS where Customer in ('" & _Customer & "') group by Department"
            ElseIf Trim(txtMerchant.Text) <> "" Then
                Sql = "select Department from OTD_SMS where Merchant in ('" & _Merchant & "') group by Department"
            Else
                Sql = "select Department from OTD_SMS group by Department"
            End If
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            I = 0

            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                'For Each DTRow1 As DataRow In M01.Tables(0).Rows
                newRow("##") = False
                newRow("Department") = M01.Tables(0).Rows(I)("Department")

                c_dataCustomer1.Rows.Add(newRow)
                I = I + 1

            Next
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

    Function Load_Merch()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Dim I As Integer

        Try

            If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                Sql = "select Merchant from OTD_SMS where Customer in ('" & _Customer & "') and Department in ('" & _Department & "')   group by Merchant"
            ElseIf Trim(txtCustomer.Text) <> "" Then
                Sql = "select Merchant from OTD_SMS where Customer in ('" & _Customer & "') group by Merchant"
            ElseIf Trim(txtDepartment.Text) <> "" Then
                Sql = "select Merchant from OTD_SMS where Department in ('" & _Department & "') group by Merchant"
            Else
                Sql = "select Merchant from OTD_SMS group by Merchant"
            End If


            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            I = 0

            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                'For Each DTRow1 As DataRow In M01.Tables(0).Rows
                newRow("##") = False
                newRow("Merchant") = M01.Tables(0).Rows(I)("Merchant")

                c_dataCustomer1.Rows.Add(newRow)
                I = I + 1

            Next
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

    Private Sub UltraButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton3.Click
        Dim i As Integer
        UltraGrid1.Visible = False
        UltraGrid3.Visible = False
        _Customer = ""
        If UltraGrid2.Visible = False Then
            Call Load_Gride()
            Call Load_Customer()
            UltraGrid2.Visible = True
        Else
            txtCustomer.Text = ""
            For Each uRow As UltraGridRow In UltraGrid2.Rows
                If UltraGrid2.Rows(i).Cells(0).Value = True Then
                    If Trim(txtCustomer.Text) <> "" Then
                        txtCustomer.Text = txtCustomer.Text & "," & UltraGrid2.Rows(i).Cells(1).Value
                        _Customer = _Customer & "','" & UltraGrid2.Rows(i).Cells(1).Value
                    Else
                        txtCustomer.Text = UltraGrid2.Rows(i).Cells(1).Value
                        _Customer = UltraGrid2.Rows(i).Cells(1).Value


                    End If
                End If
                i = i + 1
            Next
            UltraGrid2.Visible = False
        End If
    End Sub

    Private Sub UltraButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton4.Click
        Dim i As Integer
        UltraGrid2.Visible = False
        UltraGrid3.Visible = False
        _Department = ""
        If UltraGrid1.Visible = False Then
            Call Load_GrideDep()
            Call Load_Department()
            UltraGrid1.Visible = True
        Else
            txtDepartment.Text = ""
            For Each uRow As UltraGridRow In UltraGrid1.Rows
                If UltraGrid1.Rows(i).Cells(0).Value = True Then
                    If txtDepartment.Text <> "" Then
                        txtDepartment.Text = txtDepartment.Text & "," & UltraGrid1.Rows(i).Cells(1).Value
                        _Department = _Department & "','" & UltraGrid1.Rows(i).Cells(1).Value
                    Else
                        txtDepartment.Text = UltraGrid1.Rows(i).Cells(1).Value
                        _Department = UltraGrid1.Rows(i).Cells(1).Value
                    End If
                End If
                i = i + 1
            Next
            UltraGrid1.Visible = False
        End If
    End Sub

    Private Sub UltraButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton5.Click
        Dim i As Integer
        UltraGrid2.Visible = False
        UltraGrid1.Visible = False
        _Merchant = ""
        If UltraGrid3.Visible = False Then
            Call Load_GrideMerch()
            Call Load_Merch()
            UltraGrid3.Visible = True
        Else
            txtMerchant.Text = ""
            For Each uRow As UltraGridRow In UltraGrid3.Rows
                If UltraGrid3.Rows(i).Cells(0).Value = True Then
                    If txtMerchant.Text <> "" Then
                        txtMerchant.Text = txtMerchant.Text & "," & UltraGrid3.Rows(i).Cells(1).Value
                        _Merchant = _Merchant & "','" & UltraGrid3.Rows(i).Cells(1).Value
                    Else
                        txtMerchant.Text = UltraGrid3.Rows(i).Cells(1).Value
                        _Merchant = UltraGrid3.Rows(i).Cells(1).Value
                    End If
                End If
                i = i + 1
            Next
            UltraGrid3.Visible = False
        End If
    End Sub

    Private Sub chkCus_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkCus.CheckedChanged
        If chkCus.Checked = True Then
            UltraButton3.Enabled = True
        Else
            UltraButton3.Enabled = False
        End If
    End Sub

    Private Sub chkDep_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkDep.CheckedChanged
        If chkDep.Checked = True Then
            UltraButton4.Enabled = True
        Else
            UltraButton4.Enabled = False
        End If
    End Sub

    Private Sub chkMerch_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkMerch.CheckedChanged
        If chkMerch.Checked = True Then
            UltraButton5.Enabled = True
        Else
            UltraButton5.Enabled = False
        End If
    End Sub

    Function Load_Main_Gride()
        Dim _Yesterday As String
        Dim _Todayupdate As String

        _Yesterday = Today.AddDays(-1)
        _Yesterday = _Yesterday & " Update"
        _Todayupdate = Today & " Update"

        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTable_Wip(_Yesterday, _Todayupdate)
        UltraGrid4.DataSource = c_dataCustomer1
        With UltraGrid4
            .DisplayLayout.Bands(0).Columns(0).Width = 60
            .DisplayLayout.Bands(0).Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 40
            .DisplayLayout.Bands(0).Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 70
            .DisplayLayout.Bands(0).Columns(4).Width = 70
            .DisplayLayout.Bands(0).Columns(5).Width = 70
            .DisplayLayout.Bands(0).Columns(6).Width = 70
            .DisplayLayout.Bands(0).Columns(7).Width = 70
            .DisplayLayout.Bands(0).Columns(8).Width = 70

            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(3).Width = 150
            .DisplayLayout.Bands(0).Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(8).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(11).Width = 80
            .DisplayLayout.Bands(0).Columns(12).Width = 120
        End With
    End Function
    
    Private Sub frmUpdate_Status_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtFromDate.Text = Today
        txtTodate.Text = Today
        Call Load_Main_Gride()
        Call Load_PO()
        Call Load_Quality()
        ' Call Load_Customer()
        'Call Load_Merchant()
        Me.BindUltraDropDown2()
        Me.BindUltraDropDown1()
        'Call Load_Department()
        Call Load_Gride()
        Call Load_Customer()

        Call Load_GrideDep()
        Call Load_Department()

        Call Load_GrideMerch()
        Call Load_Merch()

        txtCustomer.ReadOnly = True
        txtDepartment.ReadOnly = True
        txtMerchant.ReadOnly = True

     
    End Sub

    Private Sub cboPO_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboPO.AfterCloseUp
        
    End Sub

    Function Load_Data_To_Gride()
        Dim Sql As String
        Dim M01 As DataSet
        Dim T01 As DataSet
        Dim _OrderQty As Double
        Dim I As Integer
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim _Yesterday As String
        Dim _Todayupdate As String
        Dim _delivaryDate As Date

        _Yesterday = Today.AddDays(-1)
        _Yesterday = _Yesterday & " Update"

        If Trim(cboPO.Text) = "To be planned" Then
            Dim _FROMDATE As Date
            Dim _TODATE As Date
            Dim _OTD As DataSet

            _FROMDATE = Month(txtFromDate.Text) & "/" & Microsoft.VisualBasic.Day(txtFromDate.Text) & "/" & Year(txtFromDate.Text)
            _TODATE = Month(txtTodate.Text) & "/" & Microsoft.VisualBasic.Day(txtTodate.Text) & "/" & Year(txtTodate.Text)

            If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then

                Sql = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
              "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,s.Customer " & _
             " FROM ZPP_DEL z INNER JOIN  M01Sales_Order_SAP ON Z.Sales_Order=M01Sales_Order AND M01Line_Item=Z.Line_Item " & _
            "INNER JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
            "WHERE Z.WIP_Loc='To Be planned' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' and s.Department in ('" & _Department & "') and s.Customer in ('" & _Customer & "') and s.Merchant in ('" & _Merchant & "')"

            ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then

                Sql = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
              "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,s.Customer " & _
             " FROM ZPP_DEL z INNER JOIN  M01Sales_Order_SAP ON Z.Sales_Order=M01Sales_Order AND M01Line_Item=Z.Line_Item " & _
            "INNER JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
            "WHERE Z.WIP_Loc='To Be planned' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' and s.Department in ('" & _Department & "') and s.Customer in ('" & _Customer & "')"
            ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then

                Sql = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                  "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,s.Customer " & _
                 " FROM ZPP_DEL z INNER JOIN  M01Sales_Order_SAP ON Z.Sales_Order=M01Sales_Order AND M01Line_Item=Z.Line_Item " & _
                "INNER JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                "WHERE Z.WIP_Loc='To Be planned' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' and s.Merchant in ('" & _Merchant & "') and s.Customer in ('" & _Customer & "')"
            ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then

                Sql = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                  "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,s.Customer " & _
                 " FROM ZPP_DEL z INNER JOIN  M01Sales_Order_SAP ON Z.Sales_Order=M01Sales_Order AND M01Line_Item=Z.Line_Item " & _
                "INNER JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                "WHERE Z.WIP_Loc='To Be planned' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' and s.Merchant in ('" & _Merchant & "') and s.Department in ('" & _Department & "')"
            ElseIf Trim(txtCustomer.Text) <> "" Then

                Sql = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                        "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,s.Customer " & _
                       " FROM ZPP_DEL z INNER JOIN  M01Sales_Order_SAP ON Z.Sales_Order=M01Sales_Order AND M01Line_Item=Z.Line_Item " & _
                      "INNER JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                      "WHERE Z.WIP_Loc='To Be planned' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' and s.Customer in ('" & _Customer & "')"
            ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                If cboQuality.Text <> "" Then
                    If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                        Sql = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                             "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,s.Customer " & _
                            " FROM ZPP_DEL z INNER JOIN  M01Sales_Order_SAP ON Z.Sales_Order=M01Sales_Order AND M01Line_Item=Z.Line_Item " & _
                           "INNER JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                           "WHERE Z.WIP_Loc='To Be planned' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' and s.Department in ('" & _Department & "') and left(z.Material_Dis,8)='" & cboQuality.Text & "'"
                    Else
                        Sql = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                  "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,s.Customer " & _
                 " FROM ZPP_DEL z INNER JOIN  M01Sales_Order_SAP ON Z.Sales_Order=M01Sales_Order AND M01Line_Item=Z.Line_Item " & _
                "INNER JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                "WHERE Z.WIP_Loc='To Be planned' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' and s.Department in ('" & _Department & "') and left(z.Material_Dis,5)='" & cboQuality.Text & "'"
                    End If

                Else
                    Sql = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                              "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,s.Customer " & _
                             " FROM ZPP_DEL z INNER JOIN  M01Sales_Order_SAP ON Z.Sales_Order=M01Sales_Order AND M01Line_Item=Z.Line_Item " & _
                            "INNER JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                            "WHERE Z.WIP_Loc='To Be planned' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' and s.Department in ('" & _Department & "')"

                End If
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    If Trim(cboQuality.Text) <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            Sql = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                             "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,s.Customer " & _
                            " FROM ZPP_DEL z INNER JOIN  M01Sales_Order_SAP ON Z.Sales_Order=M01Sales_Order AND M01Line_Item=Z.Line_Item " & _
                           "INNER JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                           "WHERE Z.WIP_Loc='To Be planned' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' and s.Merchant in ('" & _Merchant & "') and left(z.Material_Dis,8) ='" & cboQuality.Text & "'"
                        Else
                            Sql = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                            "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,s.Customer " & _
                           " FROM ZPP_DEL z INNER JOIN  M01Sales_Order_SAP ON Z.Sales_Order=M01Sales_Order AND M01Line_Item=Z.Line_Item " & _
                          "INNER JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                          "WHERE Z.WIP_Loc='To Be planned' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' and s.Merchant in ('" & _Merchant & "') and left(z.Material_Dis,5) ='" & cboQuality.Text & "'"
                        End If
                    Else
                        Sql = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                              "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,s.Customer " & _
                             " FROM ZPP_DEL z INNER JOIN  M01Sales_Order_SAP ON Z.Sales_Order=M01Sales_Order AND M01Line_Item=Z.Line_Item " & _
                            "INNER JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                            "WHERE Z.WIP_Loc='To Be planned' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' and s.Merchant in ('" & _Merchant & "')"
                    End If
                Else
                    If Trim(cboQuality.Text) <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            Sql = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                              "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,s.Customer " & _
                             " FROM ZPP_DEL z INNER JOIN  M01Sales_Order_SAP ON Z.Sales_Order=M01Sales_Order AND M01Line_Item=Z.Line_Item " & _
                            "INNER JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                            "WHERE Z.WIP_Loc='To Be planned' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' and left(z.Material_Dis,8)='" & cboQuality.Text & "'"
                        Else
                            Sql = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                            "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,s.Customer " & _
                           " FROM ZPP_DEL z INNER JOIN  M01Sales_Order_SAP ON Z.Sales_Order=M01Sales_Order AND M01Line_Item=Z.Line_Item " & _
                          "INNER JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                          "WHERE Z.WIP_Loc='To Be planned' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' and left(z.Material_Dis,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        Sql = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                              "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,s.Customer " & _
                             " FROM ZPP_DEL z INNER JOIN  M01Sales_Order_SAP ON Z.Sales_Order=M01Sales_Order AND M01Line_Item=Z.Line_Item " & _
                            "INNER JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                            "WHERE Z.WIP_Loc='To Be planned' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "'"
                    End If
                End If

                T01 = DBEngin.ExecuteDataset(con, Nothing, Sql)

                For Each DTRow3 As DataRow In T01.Tables(0).Rows
                    Dim newRow As DataRow = c_dataCustomer1.NewRow
                    Dim _Tollaranz As Double

                    'For Each DTRow1 As DataRow In M01.Tables(0).Rows
                    _Tollaranz = 0
                    _Tollaranz = T01.Tables(0).Rows(I)("M01SO_Qty") * (T01.Tables(0).Rows(I)("M01Cus_Tol_Min") / 100)
                    If _Tollaranz < CDbl(T01.Tables(0).Rows(I)("Order_Qty_mtr")) Then
                        Dim _Material As String
                        _Material = T01.Tables(0).Rows(I)("Material")
                        _Material = Microsoft.VisualBasic.Right(_Material, 7)
                        _Material = Microsoft.VisualBasic.Left(_Material, 2) & "-" & Microsoft.VisualBasic.Right(_Material, 5)
                        newRow("Sales Order") = T01.Tables(0).Rows(I)("Sales_Order")
                        newRow("Line Item") = T01.Tables(0).Rows(I)("Line_Item")
                        newRow("Material") = _Material
                        newRow("Description") = T01.Tables(0).Rows(I)("Material_Dis")
                        newRow("Delivary Date") = Month(T01.Tables(0).Rows(I)("Delivery_Date")) & "/" & Microsoft.VisualBasic.Day(T01.Tables(0).Rows(I)("Delivery_Date")) & "/" & Year(T01.Tables(0).Rows(I)("Delivery_Date"))
                    _delivaryDate = Month(T01.Tables(0).Rows(I)("Delivery_Date")) & "/" & Microsoft.VisualBasic.Day(T01.Tables(0).Rows(I)("Delivery_Date")) & "/" & Year(T01.Tables(0).Rows(I)("Delivery_Date"))
                    newRow("Batch No") = T01.Tables(0).Rows(I)("Product_Order")
                        newRow("Batch Qty (Kg)") = T01.Tables(0).Rows(I)("Order_Qty_Kg")
                        newRow("Batch Qty (Mtr)") = T01.Tables(0).Rows(I)("Order_Qty_mtr")
                        newRow("No.Of.Dys In S/L") = T01.Tables(0).Rows(I)("No_Day_Same_Opp")
                        newRow("NC Comment") = T01.Tables(0).Rows(I)("NC_Comment")
                        Sql = "select T07Posible_Date,T07Comment from T07OTD_Comment where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07Material='" & _Material & "' order by T07Date DESC"
                        _OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                        If isValidDataset(_OTD) Then
                            newRow("Possible Devivary Date") = T01.Tables(0).Rows(I)("T07Posible_Date")
                            newRow("Reason") = T01.Tables(0).Rows(I)("T07Reason")
                            newRow(_Yesterday) = T01.Tables(0).Rows(I)("T07Comment")
                    End If
                  
                    newRow("Week") = DatePart(DateInterval.WeekOfYear, _delivaryDate)
                        c_dataCustomer1.Rows.Add(newRow)
                    End If

                    ' If Not x = M01.Tables(0).Rows.Count - 1 Then
                    ' rsT24RecHeader.MoveNext()

                    I = I + 1
                Next

                ElseIf Trim(cboPO.Text) = "Aw Dyed Yarn" Then
                    Dim _FROMDATE As Date
                    Dim _TODATE As Date
                    Dim _OTD As DataSet

                    _FROMDATE = Month(txtFromDate.Text) & "/" & Microsoft.VisualBasic.Day(txtFromDate.Text) & "/" & Year(txtFromDate.Text)
                    _TODATE = Month(txtTodate.Text) & "/" & Microsoft.VisualBasic.Day(txtTodate.Text) & "/" & Year(txtTodate.Text)

                    If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then

                        Sql = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,s.Customer " & _
                     " FROM ZPP_DEL z INNER JOIN  M01Sales_Order_SAP ON Z.Sales_Order=M01Sales_Order AND M01Line_Item=Z.Line_Item " & _
                    "INNER JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                    "WHERE Z.WIP_Loc='Aw Dyed Yarn' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' and s.Department in ('" & _Department & "') and s.Customer in ('" & _Customer & "') and s.Merchant in ('" & _Merchant & "')"

                    ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then

                        Sql = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,s.Customer " & _
                     " FROM ZPP_DEL z INNER JOIN  M01Sales_Order_SAP ON Z.Sales_Order=M01Sales_Order AND M01Line_Item=Z.Line_Item " & _
                    "INNER JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                    "WHERE Z.WIP_Loc='Aw Dyed Yarn' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' and s.Department in ('" & _Department & "') and s.Customer in ('" & _Customer & "')"
                    ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then

                        Sql = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                          "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,s.Customer " & _
                         " FROM ZPP_DEL z INNER JOIN  M01Sales_Order_SAP ON Z.Sales_Order=M01Sales_Order AND M01Line_Item=Z.Line_Item " & _
                        "INNER JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                        "WHERE Z.WIP_Loc='Aw Dyed Yarn' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' and s.Merchant in ('" & _Merchant & "') and s.Customer in ('" & _Customer & "')"
                    ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then

                        Sql = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                          "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,s.Customer " & _
                         " FROM ZPP_DEL z INNER JOIN  M01Sales_Order_SAP ON Z.Sales_Order=M01Sales_Order AND M01Line_Item=Z.Line_Item " & _
                        "INNER JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                        "WHERE Z.WIP_Loc='Aw Dyed Yarn' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' and s.Merchant in ('" & _Merchant & "') and s.Department in ('" & _Department & "')"
                    ElseIf Trim(txtCustomer.Text) <> "" Then

                        Sql = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                                "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,s.Customer " & _
                               " FROM ZPP_DEL z INNER JOIN  M01Sales_Order_SAP ON Z.Sales_Order=M01Sales_Order AND M01Line_Item=Z.Line_Item " & _
                              "INNER JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                              "WHERE Z.WIP_Loc='Aw Dyed Yarn' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' and s.Customer in ('" & _Customer & "')"
                    ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                        Sql = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                                  "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,s.Customer " & _
                                 " FROM ZPP_DEL z INNER JOIN  M01Sales_Order_SAP ON Z.Sales_Order=M01Sales_Order AND M01Line_Item=Z.Line_Item " & _
                                "INNER JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                                "WHERE Z.WIP_Loc='Aw Dyed Yarn' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' and s.Department in ('" & _Department & "')"
                    ElseIf Trim(txtMerchant.Text) <> "" Then

                        Sql = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                              "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,s.Customer " & _
                             " FROM ZPP_DEL z INNER JOIN  M01Sales_Order_SAP ON Z.Sales_Order=M01Sales_Order AND M01Line_Item=Z.Line_Item " & _
                            "INNER JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                            "WHERE Z.WIP_Loc='Aw Dyed Yarn' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' and s.Merchant in ('" & _Merchant & "')"
                    Else
                        Sql = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                              "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,s.Customer " & _
                             " FROM ZPP_DEL z INNER JOIN  M01Sales_Order_SAP ON Z.Sales_Order=M01Sales_Order AND M01Line_Item=Z.Line_Item " & _
                            "INNER JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                            "WHERE Z.WIP_Loc='Aw Dyed Yarn' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "'"
                    End If
                    ' Dim _Last As Integer

                    T01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    For Each DTRow3 As DataRow In T01.Tables(0).Rows
                        Dim newRow As DataRow = c_dataCustomer1.NewRow

                        Dim _Material As String
                        _Material = T01.Tables(0).Rows(I)("Material")
                        _Material = Microsoft.VisualBasic.Right(_Material, 7)
                        _Material = Microsoft.VisualBasic.Left(_Material, 2) & "-" & Microsoft.VisualBasic.Right(_Material, 5)
                        newRow("Sales Order") = T01.Tables(0).Rows(I)("Sales_Order")
                        newRow("Line Item") = T01.Tables(0).Rows(I)("Line_Item")
                        newRow("Material") = _Material
                        newRow("Description") = T01.Tables(0).Rows(I)("Material_Dis")
                        newRow("Delivary Date") = Month(T01.Tables(0).Rows(I)("Delivery_Date")) & "/" & Microsoft.VisualBasic.Day(T01.Tables(0).Rows(I)("Delivery_Date")) & "/" & Year(T01.Tables(0).Rows(I)("Delivery_Date"))
                        newRow("Batch No") = T01.Tables(0).Rows(I)("Product_Order")
                        newRow("Batch Qty (Kg)") = T01.Tables(0).Rows(I)("Order_Qty_Kg")
                        newRow("Batch Qty (Mtr)") = T01.Tables(0).Rows(I)("Order_Qty_mtr")
                        newRow("No.Of.Dys In S/L") = T01.Tables(0).Rows(I)("No_Day_Same_Opp")
                        newRow("NC Comment") = T01.Tables(0).Rows(I)("NC_Comment")
                        Sql = "select T07Posible_Date,T07Comment from T07OTD_Comment where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07Material='" & _Material & "'"
                        _OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                        If isValidDataset(_OTD) Then
                            newRow("Possible Devivary Date") = T01.Tables(0).Rows(I)("T07Posible_Date")
                            newRow("Reason") = T01.Tables(0).Rows(I)("T07Comment")
                        End If
                        c_dataCustomer1.Rows.Add(newRow)


                        ' If Not x = M01.Tables(0).Rows.Count - 1 Then
                        ' rsT24RecHeader.MoveNext()

                        I = I + 1
                    Next
                    '=================================================================================
                ElseIf Trim(cboPO.Text) = "AW Greige" Then
                    Dim _FROMDATE As Date
                    Dim _TODATE As Date
                    Dim _OTD As DataSet

                    _FROMDATE = Month(txtFromDate.Text) & "/" & Microsoft.VisualBasic.Day(txtFromDate.Text) & "/" & Year(txtFromDate.Text)
                    _TODATE = Month(txtTodate.Text) & "/" & Microsoft.VisualBasic.Day(txtTodate.Text) & "/" & Year(txtTodate.Text)

                    If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then

                        Sql = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,s.Customer " & _
                     " FROM ZPP_DEL z INNER JOIN  M01Sales_Order_SAP ON Z.Sales_Order=M01Sales_Order AND M01Line_Item=Z.Line_Item " & _
                    "INNER JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                    "WHERE Z.WIP_Loc='AW Greige' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' and s.Department in ('" & _Department & "') and s.Customer in ('" & _Customer & "') and s.Merchant in ('" & _Merchant & "')"

                    ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then

                        Sql = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,s.Customer " & _
                     " FROM ZPP_DEL z INNER JOIN  M01Sales_Order_SAP ON Z.Sales_Order=M01Sales_Order AND M01Line_Item=Z.Line_Item " & _
                    "INNER JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                    "WHERE Z.WIP_Loc='AW Greige' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' and s.Department in ('" & _Department & "') and s.Customer in ('" & _Customer & "')"
                    ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then

                        Sql = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                          "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,s.Customer " & _
                         " FROM ZPP_DEL z INNER JOIN  M01Sales_Order_SAP ON Z.Sales_Order=M01Sales_Order AND M01Line_Item=Z.Line_Item " & _
                        "INNER JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                        "WHERE Z.WIP_Loc='AW Greige' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' and s.Merchant in ('" & _Merchant & "') and s.Customer in ('" & _Customer & "')"
                    ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then

                        Sql = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                          "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,s.Customer " & _
                         " FROM ZPP_DEL z INNER JOIN  M01Sales_Order_SAP ON Z.Sales_Order=M01Sales_Order AND M01Line_Item=Z.Line_Item " & _
                        "INNER JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                        "WHERE Z.WIP_Loc='AW Greige' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' and s.Merchant in ('" & _Merchant & "') and s.Department in ('" & _Department & "')"
                    ElseIf Trim(txtCustomer.Text) <> "" Then

                        Sql = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                                "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,s.Customer " & _
                               " FROM ZPP_DEL z INNER JOIN  M01Sales_Order_SAP ON Z.Sales_Order=M01Sales_Order AND M01Line_Item=Z.Line_Item " & _
                              "INNER JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                              "WHERE Z.WIP_Loc='AW Greige' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' and s.Customer in ('" & _Customer & "')"
                    ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                        Sql = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                                  "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,s.Customer " & _
                                 " FROM ZPP_DEL z INNER JOIN  M01Sales_Order_SAP ON Z.Sales_Order=M01Sales_Order AND M01Line_Item=Z.Line_Item " & _
                                "INNER JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                                "WHERE Z.WIP_Loc='AW Greige' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' and s.Department in ('" & _Department & "')"
                    ElseIf Trim(txtMerchant.Text) <> "" Then

                        Sql = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                              "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,s.Customer " & _
                             " FROM ZPP_DEL z INNER JOIN  M01Sales_Order_SAP ON Z.Sales_Order=M01Sales_Order AND M01Line_Item=Z.Line_Item " & _
                            "INNER JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                            "WHERE Z.WIP_Loc='AW Greige' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' and s.Merchant in ('" & _Merchant & "')"
                    Else
                        Sql = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                              "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,s.Customer " & _
                             " FROM ZPP_DEL z INNER JOIN  M01Sales_Order_SAP ON Z.Sales_Order=M01Sales_Order AND M01Line_Item=Z.Line_Item " & _
                            "INNER JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                            "WHERE Z.WIP_Loc='AW Greige' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "'"
                    End If
                    T01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    For Each DTRow3 As DataRow In T01.Tables(0).Rows
                        Dim newRow As DataRow = c_dataCustomer1.NewRow

                        Dim _Material As String
                        _Material = T01.Tables(0).Rows(I)("Material")
                        _Material = Microsoft.VisualBasic.Right(_Material, 7)
                        _Material = Microsoft.VisualBasic.Left(_Material, 2) & "-" & Microsoft.VisualBasic.Right(_Material, 5)
                        newRow("Sales Order") = T01.Tables(0).Rows(I)("Sales_Order")
                        newRow("Line Item") = T01.Tables(0).Rows(I)("Line_Item")
                        newRow("Material") = _Material
                        newRow("Description") = T01.Tables(0).Rows(I)("Material_Dis")
                        newRow("Delivary Date") = Month(T01.Tables(0).Rows(I)("Delivery_Date")) & "/" & Microsoft.VisualBasic.Day(T01.Tables(0).Rows(I)("Delivery_Date")) & "/" & Year(T01.Tables(0).Rows(I)("Delivery_Date"))
                        newRow("Batch No") = T01.Tables(0).Rows(I)("Product_Order")
                        newRow("Batch Qty (Kg)") = T01.Tables(0).Rows(I)("Order_Qty_Kg")
                        newRow("Batch Qty (Mtr)") = T01.Tables(0).Rows(I)("Order_Qty_mtr")
                        newRow("No.Of.Dys In S/L") = T01.Tables(0).Rows(I)("No_Day_Same_Opp")
                        newRow("NC Comment") = T01.Tables(0).Rows(I)("NC_Comment")
                        Sql = "select T07Posible_Date,T07Comment from T07OTD_Comment where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07Material='" & _Material & "'"
                        _OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                        If isValidDataset(_OTD) Then
                            newRow("Possible Devivary Date") = T01.Tables(0).Rows(I)("T07Posible_Date")
                            newRow("Reason") = T01.Tables(0).Rows(I)("T07Comment")
                        End If
                        c_dataCustomer1.Rows.Add(newRow)


                        ' If Not x = M01.Tables(0).Rows.Count - 1 Then
                        ' rsT24RecHeader.MoveNext()

                        I = I + 1
                    Next
                ElseIf Trim(cboPO.Text) = "AW Greige" Then

                End If
    End Function


    Function Load_Data_To_GrideNew()
        Dim Sql As String
        Dim M01 As DataSet
        Dim T01 As DataSet
        Dim _OrderQty As Double
        Dim qcType As String
        Dim vcCode As String
        Dim vcWharer As String
        Dim T03 As DataSet
        Dim _OTDComm As DataSet

        Dim _QtyKg As Double
        Dim _QtyMtr As Double

        Dim I As Integer
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim _Yesterday As String
        Dim _Todayupdate As String
        Dim _delivaryDT As Date
        Dim _YesDate As Date

        Dim _Rowcount As Integer
        Try
            _Yesterday = Today.AddDays(-1)
            _Yesterday = _Yesterday & " Update"
            _Todayupdate = Today & " Update"
            Call Load_Main_Gride()
            If Trim(cboPO.Text) = "To Be Planned" Then
                Dim _FROMDATE As Date
                Dim _TODATE As Date
                Dim _OTD As DataSet

                _FROMDATE = Month(txtFromDate.Text) & "/" & Microsoft.VisualBasic.Day(txtFromDate.Text) & "/" & Year(txtFromDate.Text)
                _TODATE = Month(txtTodate.Text) & "/" & Microsoft.VisualBasic.Day(txtTodate.Text) & "/" & Year(txtTodate.Text)

                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location='To Be planned' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location='To Be planned' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and left(r.Met_Des,5)='" & cboQuality.Text & "' "
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location='To Be planned' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')"

                    End If
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then

                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location='To Be planned' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and left(r.Met_Des,8)='" & cboQuality.Text & "' "
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location='To Be planned' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and left(r.Met_Des,5)='" & cboQuality.Text & "' "
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location='To Be planned' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') "
                    End If
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location='To Be planned' and s.Customer in ('" & _Customer & "') and r.Merchant in ('" & _Merchant & "') and left(r.Met_Des,8)='" & cboQuality.Text & "'  "
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location='To Be planned' and s.Customer in ('" & _Customer & "') and r.Merchant in ('" & _Merchant & "') and left(r.Met_Des,5)='" & cboQuality.Text & "' "
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location='To Be planned' and s.Customer in ('" & _Customer & "') and r.Merchant in ('" & _Merchant & "') "
                    End If
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then

                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location='To Be planned' and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and left(r.Met_Des,8)='" & cboQuality.Text & "' "
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location='To Be planned' and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and left(r.Met_Des,5)='" & cboQuality.Text & "' "
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location='To Be planned' and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') "
                    End If
                ElseIf Trim(txtCustomer.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location='To Be planned' and s.Customer in ('" & _Customer & "') and left(r.Met_Des,8)='" & cboQuality.Text & "' "
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location='To Be planned' and s.Customer in ('" & _Customer & "') and left(r.Met_Des,5)='" & cboQuality.Text & "' "
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location='To Be planned' and s.Customer in ('" & _Customer & "')"

                    End If
                ElseIf Trim(txtDepartment.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location='To Be planned' and r.Department in ('" & _Department & "') and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location='To Be planned' and r.Department in ('" & _Department & "') and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location='To Be planned' and r.Department in ('" & _Department & "')"

                    End If
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    If Trim(cboQuality.Text) <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & _FROMDATE & "' and '" & _TODATE & "' and Location='To Be planned' and r.Merchant in ('" & _Merchant & "') and left(r.Met_Des,8)='" & cboQuality.Text & "' "
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & _FROMDATE & "' and '" & _TODATE & "' and Location='To Be planned' and r.Merchant in ('" & _Merchant & "') and left(r.Met_Des,5)='" & cboQuality.Text & "' "
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & _FROMDATE & "' and '" & _TODATE & "' and Location='To Be planned' and r.Merchant in ('" & _Merchant & "') "
                    End If
                Else
                    If Trim(cboQuality.Text) <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & _FROMDATE & "' and '" & _TODATE & "' and Location='To Be planned' and left(r.Met_Des,8)='" & cboQuality.Text & "' "
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & _FROMDATE & "' and '" & _TODATE & "' and Location='To Be planned' and left(r.Met_Des,5)='" & cboQuality.Text & "' "
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & _FROMDATE & "' and '" & _TODATE & "' and Location='To Be planned'  "

                    End If
                End If
                T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetMEETING", New SqlParameter("@cQryType", "TBP"), New SqlParameter("@vcWhereClause", vcWharer))

                'T01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                _QtyKg = 0
                _QtyMtr = 0
                I = 0
                _Rowcount = 0
                For Each DTRow3 As DataRow In T01.Tables(0).Rows
                    Dim newRow As DataRow = c_dataCustomer1.NewRow
                    Dim _Tollaranz As Double

                    If I = 69 Then
                        ' MsgBox("")
                    End If
                    'For Each DTRow1 As DataRow In M01.Tables(0).Rows
                    _OrderQty = 0
                    Sql = "select * from M01Sales_Order_SAP where M01Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and M01Line_Item='" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & "'"
                    dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(dsUser) Then
                        _OrderQty = dsUser.Tables(0).Rows(0)("M01SO_Qty")
                    End If

                    If (_OrderQty * T01.Tables(0).Rows(I)("Tollarance_MIN")) / 100 >= T01.Tables(0).Rows(I)("PRD_Qty") Then
                    Else
                        Dim _Material As String
                        _Material = T01.Tables(0).Rows(I)("Metrrial")
                        '_Material = Microsoft.VisualBasic.Right(_Material, 7)
                        '_Material = Microsoft.VisualBasic.Left(_Material, 2) & "-" & Microsoft.VisualBasic.Right(_Material, 5)
                        newRow("Sales Order") = T01.Tables(0).Rows(I)("Sales_Order")
                        newRow("Line Item") = T01.Tables(0).Rows(I)("Line_Item")
                        newRow("Material") = _Material
                        newRow("Description") = T01.Tables(0).Rows(I)("Met_Des")
                        newRow("Batch Qty (Mtr)") = T01.Tables(0).Rows(I)("PRD_Qty")
                        _QtyMtr = _QtyMtr + T01.Tables(0).Rows(I)("PRD_Qty")
                        newRow("Delivary Date") = Month(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Microsoft.VisualBasic.Day(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Year(T01.Tables(0).Rows(I)("Del_Date"))
                        _delivaryDT = Month(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Microsoft.VisualBasic.Day(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Year(T01.Tables(0).Rows(I)("Del_Date"))
                        newRow("Batch No") = T01.Tables(0).Rows(I)("Prduct_Order")
                        newRow("Customer") = T01.Tables(0).Rows(I)("Customer")

                        Sql = "select * from ZPP_DEL where Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and Line_Item='" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & "' and Delivery_Date='" & T01.Tables(0).Rows(I)("Del_Date") & "' and WIP_Loc='To Be planned' "
                        dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
                        If isValidDataset(dsUser) Then
                            newRow("Batch Qty (Kg)") = Microsoft.VisualBasic.Format(dsUser.Tables(0).Rows(0)("Order_Qty_Kg"), "#.00")
                            _QtyKg = _QtyKg + dsUser.Tables(0).Rows(0)("Order_Qty_Kg")
                            newRow("No.Of.Dys In S/L") = dsUser.Tables(0).Rows(0)("No_Day_Same_Opp")
                            newRow("NC Comment") = dsUser.Tables(0).Rows(0)("NC_Comment")
                        End If
                        Sql = "select * from T07OTD_Comment1 where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07Material='" & _Material & "' order by T07Date DESC"
                        _OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                        If isValidDataset(_OTD) Then
                            If _OTD.Tables(0).Rows(0)("T07Date") = Today Then
                                If IsDate(_OTD.Tables(0).Rows(0)("T07Posible_Date")) Then
                                    newRow("Possible Devivary Date") = _OTD.Tables(0).Rows(0)("T07Posible_Date")
                                End If
                                newRow("Reason") = _OTD.Tables(0).Rows(0)("T07Reason")
                                'newRow(_Todayupdate) = _OTD.Tables(0).Rows(0)("T07Comment")
                                newRow("LIB dep") = _OTD.Tables(0).Rows(0)("T07LIB")
                            Else

                                ' newRow("Possible Devivary Date") = _OTD.Tables(0).Rows(0)("T07Posible_Date")
                                '  newRow("Reason") = _OTD.Tables(0).Rows(0)("T07Reason")
                                'newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")
                                '   newRow("LIB dep") = _OTD.Tables(0).Rows(0)("T07LIB")
                            End If
                        End If
                        newRow("Week") = "Week " & DatePart(DateInterval.WeekOfYear, _delivaryDT)
                        c_dataCustomer1.Rows.Add(newRow)
                        _Rowcount = _Rowcount + 1
                    End If

                    ' If Not x = M01.Tables(0).Rows.Count - 1 Then
                    ' rsT24RecHeader.MoveNext()

                    I = I + 1
                Next

                'Dim newRow2 As DataRow = c_dataCustomer1.NewRow
                'newRow2("Batch Qty (Kg)") = Microsoft.VisualBasic.Format(_QtyKg, "#.00")
                'newRow2("Batch Qty (Mtr)") = Microsoft.VisualBasic.Format(_QtyMtr, "#.00")
                'c_dataCustomer1.Rows.Add(newRow2)
                'UltraGrid4.Rows(I).Cells(6).Appearance.BackColor = Color.DeepSkyBlue
                'UltraGrid4.Rows(I).Cells(7).Appearance.BackColor = Color.DeepSkyBlue



            ElseIf Trim(cboPO.Text) = "Aw Yarn" Then
                Dim _FROMDATE As Date
                Dim _TODATE As Date
                Dim _OTD As DataSet

                _FROMDATE = Month(txtFromDate.Text) & "/" & Microsoft.VisualBasic.Day(txtFromDate.Text) & "/" & Year(txtFromDate.Text)
                _TODATE = Month(txtTodate.Text) & "/" & Microsoft.VisualBasic.Day(txtTodate.Text) & "/" & Year(txtTodate.Text)

                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location='Aw yarn' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location='Aw yarn' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and left(r.Met_Des,5)='" & cboQuality.Text & "'  "
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location='Aw yarn' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') "

                    End If
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then

                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location='Aw yarn' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location='Aw yarn' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location='Aw yarn' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')"
                    End If
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location='Aw yarn' and s.Customer in ('" & _Customer & "') and r.Merchant in ('" & _Merchant & "') and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location='Aw yarn' and s.Customer in ('" & _Customer & "') and r.Merchant in ('" & _Merchant & "') and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location='Aw yarn' and s.Customer in ('" & _Customer & "') and r.Merchant in ('" & _Merchant & "')"
                    End If
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then

                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location='Aw yarn' and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location='Aw yarn' and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location='Aw yarn' and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')"
                    End If
                ElseIf Trim(txtCustomer.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location='Aw yarn' and s.Customer in ('" & _Customer & "') and left(r.Met_Des,8)='" & cboQuality.Text & "' "
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location='Aw yarn' and s.Customer in ('" & _Customer & "') and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location='Aw yarn' and s.Customer in ('" & _Customer & "')"

                    End If
                ElseIf Trim(txtDepartment.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location='Aw yarn' and r.Department in ('" & _Department & "') and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location='Aw yarn' and r.Department in ('" & _Department & "') and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location='Aw yarn' and r.Department in ('" & _Department & "')"

                    End If
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    If Trim(cboQuality.Text) <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & _FROMDATE & "' and '" & _TODATE & "' and Location='Aw yarn' and r.Merchant in ('" & _Merchant & "') and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & _FROMDATE & "' and '" & _TODATE & "' and Location='Aw yarn' and r.Merchant in ('" & _Merchant & "') and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & _FROMDATE & "' and '" & _TODATE & "' and Location='Aw yarn' and r.Merchant in ('" & _Merchant & "')"
                    End If
                Else
                    If Trim(cboQuality.Text) <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & _FROMDATE & "' and '" & _TODATE & "' and Location='Aw yarn' and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & _FROMDATE & "' and '" & _TODATE & "' and Location='Aw yarn' and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & _FROMDATE & "' and '" & _TODATE & "' and Location='Aw yarn' "

                    End If
                End If
                I = 0

                _QtyKg = 0
                _QtyMtr = 0
                ' Dim _Last As Integer
                _Rowcount = 0
                T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetMEETING", New SqlParameter("@cQryType", "TBP"), New SqlParameter("@vcWhereClause", vcWharer))
                'T01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                For Each DTRow3 As DataRow In T01.Tables(0).Rows
                    Dim newRow As DataRow = c_dataCustomer1.NewRow
                    Dim _Tollaranz As Double

                    'For Each DTRow1 As DataRow In M01.Tables(0).Rows


                    Dim _Material As String
                    _Material = T01.Tables(0).Rows(I)("Metrrial")
                    '_Material = Microsoft.VisualBasic.Right(_Material, 7)
                    '_Material = Microsoft.VisualBasic.Left(_Material, 2) & "-" & Microsoft.VisualBasic.Right(_Material, 5)
                    newRow("Sales Order") = T01.Tables(0).Rows(I)("Sales_Order")
                    newRow("Line Item") = T01.Tables(0).Rows(I)("Line_Item")
                    newRow("Material") = _Material
                    newRow("Description") = T01.Tables(0).Rows(I)("Met_Des")
                    newRow("Batch Qty (Mtr)") = T01.Tables(0).Rows(I)("PRD_Qty")
                    _QtyMtr = _QtyMtr + T01.Tables(0).Rows(I)("PRD_Qty")
                    newRow("Delivary Date") = Month(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Microsoft.VisualBasic.Day(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Year(T01.Tables(0).Rows(I)("Del_Date"))
                    _delivaryDT = Month(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Microsoft.VisualBasic.Day(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Year(T01.Tables(0).Rows(I)("Del_Date"))
                    newRow("Batch No") = T01.Tables(0).Rows(I)("Prduct_Order")
                    newRow("Customer") = T01.Tables(0).Rows(I)("Customer")


                    Sql = "select * from ZPP_DEL where Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and Line_Item='" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & "' and Delivery_Date='" & T01.Tables(0).Rows(I)("Del_Date") & "' and Product_Order='" & Trim(T01.Tables(0).Rows(I)("Prduct_Order")) & "'"
                    dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(dsUser) Then
                        newRow("Batch Qty (Kg)") = dsUser.Tables(0).Rows(0)("Order_Qty_Kg")
                        _QtyKg = _QtyKg + dsUser.Tables(0).Rows(0)("Order_Qty_Kg")
                        newRow("No.Of.Dys In S/L") = dsUser.Tables(0).Rows(0)("No_Day_Same_Opp")
                        newRow("NC Comment") = dsUser.Tables(0).Rows(0)("NC_Comment")
                        newRow("Next Operation") = dsUser.Tables(0).Rows(0)("Status")
                    End If

                    Sql = "select * from T07OTD_Comment1 where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07BatchNo='" & T01.Tables(0).Rows(I)("Prduct_Order") & "' order by T07Date DESC"
                    _OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(_OTD) Then
                        If _OTD.Tables(0).Rows(0)("T07Date") = Today Then
                            newRow("Possible Devivary Date") = _OTD.Tables(0).Rows(0)("T07Posible_Date")
                            newRow("Reason") = _OTD.Tables(0).Rows(0)("T07Reason")
                           
                            newRow(_Todayupdate) = _OTD.Tables(0).Rows(0)("T07Comment")
                            newRow("LIB dep") = _OTD.Tables(0).Rows(0)("T07LIB")

                            _YesDate = Today.AddDays(-1)
                            Sql = "select * from T07OTD_Comment1 where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07BatchNo='" & T01.Tables(0).Rows(I)("Prduct_Order") & "' and T07Date<'" & Today & "' order by T07Date DESC"
                            _OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                            If isValidDataset(_OTD) Then
                                newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")
                            End If
                        Else

                            '   newRow("Possible Devivary Date") = _OTD.Tables(0).Rows(0)("T07Posible_Date")
                            '  newRow("Reason") = _OTD.Tables(0).Rows(0)("T07Reason")
                            newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")
                            ' newRow("LIB dep") = _OTD.Tables(0).Rows(0)("T07LIB")
                        End If
                    End If
                    newRow("Week") = "Week " & DatePart(DateInterval.WeekOfYear, _delivaryDT)
                    c_dataCustomer1.Rows.Add(newRow)
                    ' End If

                    ' If Not x = M01.Tables(0).Rows.Count - 1 Then
                    ' rsT24RecHeader.MoveNext()
                    _Rowcount = _Rowcount + 1
                    I = I + 1
                Next

                'Dim newRow2 As DataRow = c_dataCustomer1.NewRow
                'newRow2("Batch Qty (Kg)") = Microsoft.VisualBasic.Format(_QtyKg, "#.00")
                'newRow2("Batch Qty (Mtr)") = Microsoft.VisualBasic.Format(_QtyMtr, "#.00")
                'c_dataCustomer1.Rows.Add(newRow2)
                'UltraGrid4.Rows(I).Cells(6).Appearance.BackColor = Color.DeepSkyBlue
                'UltraGrid4.Rows(I).Cells(7).Appearance.BackColor = Color.DeepSkyBlue

                '=================================================================================
            ElseIf Trim(cboPO.Text) = "Aw Dyed Yarn" Then
                Dim _FROMDATE As Date
                Dim _TODATE As Date
                Dim _OTD As DataSet

                _FROMDATE = Month(txtFromDate.Text) & "/" & Microsoft.VisualBasic.Day(txtFromDate.Text) & "/" & Year(txtFromDate.Text)
                _TODATE = Month(txtTodate.Text) & "/" & Microsoft.VisualBasic.Day(txtTodate.Text) & "/" & Year(txtTodate.Text)

                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location='Aw Dyed Yarn' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and left(r.Met_Des,8)='" & cboQuality.Text & "' "
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location='Aw Dyed Yarn' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and left(r.Met_Des,5)='" & cboQuality.Text & "' "
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location='Aw Dyed Yarn' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')"

                    End If
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then

                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location='Aw Dyed Yarn' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and left(r.Met_Des,8)='" & cboQuality.Text & "' "
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location='Aw Dyed Yarn' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and left(r.Met_Des,5)='" & cboQuality.Text & "' "
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location='Aw Dyed Yarn' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')"
                    End If
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location='Aw Dyed Yarn' and s.Customer in ('" & _Customer & "') and r.Merchant in ('" & _Merchant & "') and left(r.Met_Des,8)='" & cboQuality.Text & "' "
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location='Aw Dyed Yarn' and s.Customer in ('" & _Customer & "') and r.Merchant in ('" & _Merchant & "') and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location='Aw Dyed Yarn' and s.Customer in ('" & _Customer & "') and r.Merchant in ('" & _Merchant & "')"
                    End If
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then

                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location='Aw Dyed Yarn' and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location='Aw Dyed Yarn' and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location='Aw Dyed Yarn' and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')"
                    End If
                ElseIf Trim(txtCustomer.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location='Aw Dyed Yarn' and s.Customer in ('" & _Customer & "') and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location='Aw Dyed Yarn' and s.Customer in ('" & _Customer & "') and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location='Aw Dyed Yarn' and s.Customer in ('" & _Customer & "') "

                    End If
                ElseIf Trim(txtDepartment.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location='Aw Dyed Yarn' and r.Department in ('" & _Department & "') and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location='Aw Dyed Yarn' and r.Department in ('" & _Department & "') and left(r.Met_Des,5)='" & cboQuality.Text & "' "
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location='Aw Dyed Yarn' and r.Department in ('" & _Department & "')"

                    End If
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    If Trim(cboQuality.Text) <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & _FROMDATE & "' and '" & _TODATE & "' and Location='Aw Dyed Yarn' and r.Merchant in ('" & _Merchant & "') and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & _FROMDATE & "' and '" & _TODATE & "' and Location='Aw Dyed Yarn' and r.Merchant in ('" & _Merchant & "') and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & _FROMDATE & "' and '" & _TODATE & "' and Location='Aw Dyed Yarn' and r.Merchant in ('" & _Merchant & "')"
                    End If
                Else
                    If Trim(cboQuality.Text) <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & _FROMDATE & "' and '" & _TODATE & "' and Location='Aw Dyed Yarn' and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & _FROMDATE & "' and '" & _TODATE & "' and Location='Aw Dyed Yarn' and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & _FROMDATE & "' and '" & _TODATE & "' and Location='Aw Dyed Yarn'"

                    End If
                End If
                I = 0
                'T01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetMEETING", New SqlParameter("@cQryType", "TBP"), New SqlParameter("@vcWhereClause", vcWharer))
                'T01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                _QtyKg = 0
                _QtyMtr = 0
                _Rowcount = 0
                For Each DTRow3 As DataRow In T01.Tables(0).Rows
                    Dim newRow As DataRow = c_dataCustomer1.NewRow
                    Dim _Tollaranz As Double

                    'For Each DTRow1 As DataRow In M01.Tables(0).Rows


                    Dim _Material As String
                    _Material = T01.Tables(0).Rows(I)("Metrrial")
                    '_Material = Microsoft.VisualBasic.Right(_Material, 7)
                    '_Material = Microsoft.VisualBasic.Left(_Material, 2) & "-" & Microsoft.VisualBasic.Right(_Material, 5)
                    newRow("Sales Order") = T01.Tables(0).Rows(I)("Sales_Order")
                    newRow("Line Item") = T01.Tables(0).Rows(I)("Line_Item")
                    newRow("Material") = _Material
                    newRow("Description") = T01.Tables(0).Rows(I)("Met_Des")
                    newRow("Batch Qty (Mtr)") = T01.Tables(0).Rows(I)("PRD_Qty")
                    _QtyMtr = _QtyMtr + T01.Tables(0).Rows(I)("PRD_Qty")
                    newRow("Delivary Date") = Month(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Microsoft.VisualBasic.Day(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Year(T01.Tables(0).Rows(I)("Del_Date"))
                    _delivaryDT = Month(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Microsoft.VisualBasic.Day(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Year(T01.Tables(0).Rows(I)("Del_Date"))
                    newRow("Batch No") = T01.Tables(0).Rows(I)("Prduct_Order")
                    newRow("Customer") = T01.Tables(0).Rows(I)("Customer")

                    Sql = "select * from ZPP_DEL where Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and Line_Item='" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & "' and Delivery_Date='" & T01.Tables(0).Rows(I)("Del_Date") & "' and Product_Order='" & Trim(T01.Tables(0).Rows(I)("Prduct_Order")) & "'"
                    dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(dsUser) Then
                        newRow("Batch Qty (Kg)") = dsUser.Tables(0).Rows(0)("Order_Qty_Kg")
                        _QtyKg = _QtyKg + dsUser.Tables(0).Rows(0)("Order_Qty_Kg")
                        newRow("No.Of.Dys In S/L") = dsUser.Tables(0).Rows(0)("No_Day_Same_Opp")
                        newRow("NC Comment") = dsUser.Tables(0).Rows(0)("NC_Comment")
                        newRow("Next Operation") = dsUser.Tables(0).Rows(0)("Status")
                    End If


                 

                    Sql = "select * from T07OTD_Comment1 where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07Material='" & _Material & "' and T07BatchNo='" & T01.Tables(0).Rows(I)("Prduct_Order") & "' order by T07Date DESC"
                    _OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(_OTD) Then

                        If Year(_OTD.Tables(0).Rows(0)("T07Posible_Date")) = "1900" Then

                        Else
                            newRow("Possible Devivary Date") = _OTD.Tables(0).Rows(0)("T07Posible_Date")

                        End If
                        newRow("Reason") = _OTD.Tables(0).Rows(0)("T07Reason")
                        If _OTD.Tables(0).Rows(0)("T07date") = Today Then
                            newRow(_Todayupdate) = _OTD.Tables(0).Rows(0)("T07Comment")
                        Else
                            newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")
                        End If
                        newRow("LIB dep") = _OTD.Tables(0).Rows(0)("T07LIB")
                        _YesDate = Today.AddDays(-1)
                        Sql = "select * from T07OTD_Comment1 where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07BatchNo='" & T01.Tables(0).Rows(I)("Prduct_Order") & "' and T07Date<'" & Today & "' order by T07Date DESC"
                        _OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                        If isValidDataset(_OTD) Then
                            newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")
                        End If
                    Else
                        'Sql = "select * from T07OTD_Comment where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07Material='" & _Material & "'  order by T07Date DESC"
                        '_OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                        'If isValidDataset(_OTD) Then
                        '    'newRow("Possible Devivary Date") = _OTD.Tables(0).Rows(0)("T07Posible_Date")
                        '    'newRow("Reason") = _OTD.Tables(0).Rows(0)("T07Reason")
                        '    newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")
                        '    'newRow("LIB dep") = _OTD.Tables(0).Rows(0)("T07LIB")
                        'End If
                    End If

                    newRow("Week") = "Week " & DatePart(DateInterval.WeekOfYear, _delivaryDT)
                    c_dataCustomer1.Rows.Add(newRow)
                    ' End If

                    ' If Not x = M01.Tables(0).Rows.Count - 1 Then
                    ' rsT24RecHeader.MoveNext()
                    _Rowcount = _Rowcount + 1
                    I = I + 1
                Next

                'Dim newRow2 As DataRow = c_dataCustomer1.NewRow
                'newRow2("Batch Qty (Kg)") = Microsoft.VisualBasic.Format(_QtyKg, "#.00")
                'newRow2("Batch Qty (Mtr)") = Microsoft.VisualBasic.Format(_QtyMtr, "#.00")
                'c_dataCustomer1.Rows.Add(newRow2)
                'UltraGrid4.Rows(I).Cells(6).Appearance.BackColor = Color.DeepSkyBlue
                'UltraGrid4.Rows(I).Cells(7).Appearance.BackColor = Color.DeepSkyBlue

            ElseIf Trim(cboPO.Text) = "Aw Greige" Then
                Dim _FROMDATE As Date
                Dim _TODATE As Date
                Dim _OTD As DataSet

                _FROMDATE = Month(txtFromDate.Text) & "/" & Microsoft.VisualBasic.Day(txtFromDate.Text) & "/" & Year(txtFromDate.Text)
                _TODATE = Month(txtTodate.Text) & "/" & Microsoft.VisualBasic.Day(txtTodate.Text) & "/" & Year(txtTodate.Text)

                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Awaiting_Grie='Awaiting Greige' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Awaiting_Grie='Awaiting Greige' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Awaiting_Grie='Awaiting Greige' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')"

                    End If
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then

                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Awaiting_Grie='Awaiting Greige' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Awaiting_Grie='Awaiting Greige' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Awaiting_Grie='Awaiting Greige' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')  order by r.Del_Date"
                    End If
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Awaiting_Grie='Awaiting Greige' and s.Customer in ('" & _Customer & "') and r.Merchant in ('" & _Merchant & "') and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Awaiting_Grie='Awaiting Greige' and s.Customer in ('" & _Customer & "') and r.Merchant in ('" & _Merchant & "') and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Awaiting_Grie='Awaiting Greige' and s.Customer in ('" & _Customer & "') and r.Merchant in ('" & _Merchant & "')"
                    End If
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then

                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Awaiting_Grie='Awaiting Greige' and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Awaiting_Grie='Awaiting Greige' and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Awaiting_Grie='Awaiting Greige' and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')"
                    End If
                ElseIf Trim(txtCustomer.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Awaiting_Grie='Awaiting Greige' and s.Customer in ('" & _Customer & "') and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Awaiting_Grie='Awaiting Greige' and s.Customer in ('" & _Customer & "') and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Awaiting_Grie='Awaiting Greige' and s.Customer in ('" & _Customer & "')"

                    End If
                ElseIf Trim(txtDepartment.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Awaiting_Grie='Awaiting Greige' and r.Department in ('" & _Department & "') and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Awaiting_Grie='Awaiting Greige' and r.Department in ('" & _Department & "') and left(r.Met_Des,5)='" & cboQuality.Text & "' "
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Awaiting_Grie='Awaiting Greige' and r.Department in ('" & _Department & "') "

                    End If
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    If Trim(cboQuality.Text) <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & _FROMDATE & "' and '" & _TODATE & "' and Awaiting_Grie='Awaiting Greige' and r.Merchant in ('" & _Merchant & "') and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & _FROMDATE & "' and '" & _TODATE & "' and Awaiting_Grie='Awaiting Greige' and r.Merchant in ('" & _Merchant & "') and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & _FROMDATE & "' and '" & _TODATE & "' and Awaiting_Grie='Awaiting Greige' and r.Merchant in ('" & _Merchant & "')"
                    End If
                Else
                    If Trim(cboQuality.Text) <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & _FROMDATE & "' and '" & _TODATE & "' and Awaiting_Grie='Awaiting Greige' and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & _FROMDATE & "' and '" & _TODATE & "' and Awaiting_Grie='Awaiting Greige' and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & _FROMDATE & "' and '" & _TODATE & "' and Awaiting_Grie='Awaiting Greige'  "

                    End If
                End If
                I = 0
                'T01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetMEETING", New SqlParameter("@cQryType", "GRG"), New SqlParameter("@vcWhereClause", vcWharer))
                'T01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                _Rowcount = 0
                _QtyKg = 0
                _QtyMtr = 0
                For Each DTRow3 As DataRow In T01.Tables(0).Rows
                    Dim newRow As DataRow = c_dataCustomer1.NewRow
                    Dim _Tollaranz As Double

                    'For Each DTRow1 As DataRow In M01.Tables(0).Rows


                    Dim _Material As String
                    _Material = T01.Tables(0).Rows(I)("Metrrial")
                    '_Material = Microsoft.VisualBasic.Right(_Material, 7)
                    '_Material = Microsoft.VisualBasic.Left(_Material, 2) & "-" & Microsoft.VisualBasic.Right(_Material, 5)
                    newRow("Sales Order") = T01.Tables(0).Rows(I)("Sales_Order")
                    newRow("Line Item") = T01.Tables(0).Rows(I)("Line_Item")
                    newRow("Material") = _Material
                    newRow("Description") = T01.Tables(0).Rows(I)("Met_Des")
                    newRow("Batch Qty (Mtr)") = Microsoft.VisualBasic.Format(T01.Tables(0).Rows(I)("PRD_Qty"), "#.00")
                    newRow("Delivary Date") = Month(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Microsoft.VisualBasic.Day(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Year(T01.Tables(0).Rows(I)("Del_Date"))
                    _delivaryDT = Month(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Microsoft.VisualBasic.Day(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Year(T01.Tables(0).Rows(I)("Del_Date"))
                    newRow("Batch No") = T01.Tables(0).Rows(I)("Prduct_Order")
                    _QtyMtr = _QtyMtr + T01.Tables(0).Rows(I)("PRD_Qty")
                    newRow("Customer") = T01.Tables(0).Rows(I)("Customer")

                    Sql = "select * from ZPP_DEL where Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and Line_Item='" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & "' and Delivery_Date='" & T01.Tables(0).Rows(I)("Del_Date") & "' and Product_Order='" & Trim(T01.Tables(0).Rows(I)("Prduct_Order")) & "'"
                    dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(dsUser) Then
                        newRow("Batch Qty (Kg)") = Microsoft.VisualBasic.Format(dsUser.Tables(0).Rows(0)("Order_Qty_Kg"), "#.00")
                        _QtyKg = _QtyKg + dsUser.Tables(0).Rows(0)("Order_Qty_Kg")
                        newRow("No.Of.Dys In S/L") = dsUser.Tables(0).Rows(0)("No_Day_Same_Opp")
                        newRow("NC Comment") = dsUser.Tables(0).Rows(0)("NC_Comment")
                        newRow("Next Operation") = dsUser.Tables(0).Rows(0)("Status")
                    End If

                    Sql = "select * from T07OTD_Comment1 where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07Material='" & _Material & "' and T07BatchNo='" & T01.Tables(0).Rows(I)("Prduct_Order") & "' order by T07Date DESC"
                    _OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(_OTD) Then
                        If IsDate(_OTD.Tables(0).Rows(0)("T07Posible_Date")) Then
                            If Year(_OTD.Tables(0).Rows(0)("T07Posible_Date")) = "1900" Then
                            Else
                                newRow("Possible Devivary Date") = _OTD.Tables(0).Rows(0)("T07Posible_Date")

                            End If
                        End If
                        newRow("Reason") = _OTD.Tables(0).Rows(0)("T07Reason")
                        If _OTD.Tables(0).Rows(0)("T07date") = Today Then
                            newRow(_Todayupdate) = _OTD.Tables(0).Rows(0)("T07Comment")
                        Else
                            newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")
                        End If
                        newRow("LIB dep") = _OTD.Tables(0).Rows(0)("T07LIB")

                        _YesDate = Today.AddDays(-1)
                        Sql = "select * from T07OTD_Comment1 where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07BatchNo='" & T01.Tables(0).Rows(I)("Prduct_Order") & "' and T07Date<'" & Today & "' order by T07Date DESC"
                        _OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                        If isValidDataset(_OTD) Then
                            newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")
                        End If
                    Else
                        'Sql = "select * from T07OTD_Comment where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07Material='" & _Material & "'  order by T07Date DESC"
                        '_OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                        'If isValidDataset(_OTD) Then
                        '    'newRow("Possible Devivary Date") = _OTD.Tables(0).Rows(0)("T07Posible_Date")
                        '    'newRow("Reason") = _OTD.Tables(0).Rows(0)("T07Reason")
                        '    newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")
                        '    'newRow("LIB dep") = _OTD.Tables(0).Rows(0)("T07LIB")
                        'End If
                    End If

                    newRow("Week") = "Week " & DatePart(DateInterval.WeekOfYear, _delivaryDT)
                    c_dataCustomer1.Rows.Add(newRow)
                    ' End If

                    ' If Not x = M01.Tables(0).Rows.Count - 1 Then
                    ' rsT24RecHeader.MoveNext()
                    _Rowcount = _Rowcount + 1
                    I = I + 1
                Next
            ElseIf Trim(cboPO.Text) = "To be Batch Card Send" Then
                Dim _FROMDATE As Date
                Dim _TODATE As Date
                Dim _OTD As DataSet

                _FROMDATE = Month(txtFromDate.Text) & "/" & Microsoft.VisualBasic.Day(txtFromDate.Text) & "/" & Year(txtFromDate.Text)
                _TODATE = Month(txtTodate.Text) & "/" & Microsoft.VisualBasic.Day(txtTodate.Text) & "/" & Year(txtTodate.Text)

                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')"

                    End If
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then

                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = "s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = "s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = "s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')"
                    End If
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = "s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "') and left(r.Met_Des,8)='" & cboQuality.Text & "' "
                        Else
                            vcWharer = "s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "') and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = "s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "')  "
                    End If
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then

                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and  r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and left(r.Met_Des,8)='" & cboQuality.Text & "'  "
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and  r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and left(r.Met_Des,5)='" & cboQuality.Text & "'  "
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and  r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  "
                    End If
                ElseIf Trim(txtCustomer.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and s.Customer in ('" & _Customer & "') and left(r.Met_Des,8)='" & cboQuality.Text & "'  "
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and s.Customer in ('" & _Customer & "') and left(r.Met_Des,5)='" & cboQuality.Text & "'  "
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and s.Customer in ('" & _Customer & "')  "

                    End If
                ElseIf Trim(txtDepartment.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = "   s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and  r.Department in ('" & _Department & "')  and left(r.Met_Des,8)='" & cboQuality.Text & "'  "
                        Else
                            vcWharer = "   s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and  r.Department in ('" & _Department & "')  and left(r.Met_Des,5)='" & cboQuality.Text & "'  "
                        End If
                    Else
                        vcWharer = "   s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and  r.Department in ('" & _Department & "')  "
                    End If
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    If Trim(cboQuality.Text) <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = "s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and  r.Merchant in ('" & _Merchant & "') and left(r.Met_Des,8)='" & cboQuality.Text & "'  "
                        Else
                            vcWharer = "s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and  r.Merchant in ('" & _Merchant & "') and left(r.Met_Des,5)='" & cboQuality.Text & "' "
                        End If
                    Else
                        vcWharer = "s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and  r.Merchant in ('" & _Merchant & "') "
                    End If
                Else
                    If Trim(cboQuality.Text) <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and left(r.Met_Des,8)='" & cboQuality.Text & "'  "
                        Else
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and left(r.Met_Des,5)='" & cboQuality.Text & "' "
                        End If
                    Else
                        vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  "

                    End If
                End If
                I = 0
                'T01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetMEETING", New SqlParameter("@cQryType", "TOB"), New SqlParameter("@vcWhereClause", vcWharer))
                'T01 = DBEngin.ExecuteDa
                _QtyKg = 0
                _QtyMtr = 0
                _Rowcount = _Rowcount + 1
                For Each DTRow3 As DataRow In T01.Tables(0).Rows
                    Dim newRow As DataRow = c_dataCustomer1.NewRow
                    Dim _Tollaranz As Double

                    'For Each DTRow1 As DataRow In M01.Tables(0).Rows


                    Dim _Material As String
                    _Material = T01.Tables(0).Rows(I)("Metrrial")
                    '_Material = Microsoft.VisualBasic.Right(_Material, 7)
                    '_Material = Microsoft.VisualBasic.Left(_Material, 2) & "-" & Microsoft.VisualBasic.Right(_Material, 5)
                    newRow("Sales Order") = T01.Tables(0).Rows(I)("Sales_Order")
                    newRow("Line Item") = T01.Tables(0).Rows(I)("Line_Item")
                    newRow("Material") = _Material
                    newRow("Description") = T01.Tables(0).Rows(I)("Met_Des")
                    newRow("Batch Qty (Mtr)") = Microsoft.VisualBasic.Format(T01.Tables(0).Rows(I)("PRD_Qty"), "#.00")
                    _QtyMtr = _QtyMtr + T01.Tables(0).Rows(I)("PRD_Qty")
                    newRow("Delivary Date") = Month(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Microsoft.VisualBasic.Day(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Year(T01.Tables(0).Rows(I)("Del_Date"))
                    _delivaryDT = Month(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Microsoft.VisualBasic.Day(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Year(T01.Tables(0).Rows(I)("Del_Date"))

                    newRow("Batch No") = T01.Tables(0).Rows(I)("Prduct_Order")
                    newRow("Customer") = T01.Tables(0).Rows(I)("Customer")

                    Sql = "select * from ZPP_DEL where Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and Line_Item='" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & "' and Delivery_Date='" & T01.Tables(0).Rows(I)("Del_Date") & "' and Product_Order='" & Trim(T01.Tables(0).Rows(I)("Prduct_Order")) & "'"
                    dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(dsUser) Then
                        newRow("Batch Qty (Kg)") = Microsoft.VisualBasic.Format(dsUser.Tables(0).Rows(0)("Order_Qty_Kg"), "#.00")
                        _QtyKg = _QtyKg + dsUser.Tables(0).Rows(0)("Order_Qty_Kg")
                        newRow("No.Of.Dys In S/L") = dsUser.Tables(0).Rows(0)("No_Day_Same_Opp")
                        newRow("NC Comment") = dsUser.Tables(0).Rows(0)("NC_Comment")
                        newRow("Next Operation") = dsUser.Tables(0).Rows(0)("Status")
                    End If

                    Sql = "select * from T07OTD_Comment1 where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07Material='" & _Material & "' and T07BatchNo='" & T01.Tables(0).Rows(I)("Prduct_Order") & "' order by T07Date DESC"
                    _OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(_OTD) Then
                        If IsDate(_OTD.Tables(0).Rows(0)("T07Posible_Date")) Then
                            If Year(_OTD.Tables(0).Rows(0)("T07Posible_Date")) = "1900" Then
                            Else
                                newRow("Possible Devivary Date") = _OTD.Tables(0).Rows(0)("T07Posible_Date")

                            End If
                        End If

                        newRow("Reason") = _OTD.Tables(0).Rows(0)("T07Reason")
                        If _OTD.Tables(0).Rows(0)("T07date") = Today Then
                            newRow(_Todayupdate) = _OTD.Tables(0).Rows(0)("T07Comment")
                        Else
                            newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")
                        End If
                        newRow("LIB dep") = _OTD.Tables(0).Rows(0)("T07LIB")
                        _YesDate = Today.AddDays(-1)
                        Sql = "select * from T07OTD_Comment1 where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07BatchNo='" & T01.Tables(0).Rows(I)("Prduct_Order") & "' and T07Date<'" & Today & "' order by T07Date DESC"
                        _OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                        If isValidDataset(_OTD) Then
                            newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")
                        End If
                    Else
                        'Sql = "select * from T07OTD_Comment where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07Material='" & _Material & "'  order by T07Date DESC"
                        '_OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                        'If isValidDataset(_OTD) Then
                        '    'newRow("Possible Devivary Date") = _OTD.Tables(0).Rows(0)("T07Posible_Date")
                        '    'newRow("Reason") = _OTD.Tables(0).Rows(0)("T07Reason")
                        '    newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")
                        '    'newRow("LIB dep") = _OTD.Tables(0).Rows(0)("T07LIB")
                        'End If
                    End If

                    newRow("Week") = "Week " & DatePart(DateInterval.WeekOfYear, _delivaryDT)
                    c_dataCustomer1.Rows.Add(newRow)
                    ' End If

                    ' If Not x = M01.Tables(0).Rows.Count - 1 Then
                    ' rsT24RecHeader.MoveNext()
                    _Rowcount = _Rowcount + 1
                    I = I + 1
                Next
            ElseIf Trim(cboPO.Text) = "Aw Recipe" Then
                Dim _FROMDATE As Date
                Dim _TODATE As Date
                Dim _OTD As DataSet

                _FROMDATE = Month(txtFromDate.Text) & "/" & Microsoft.VisualBasic.Day(txtFromDate.Text) & "/" & Year(txtFromDate.Text)
                _TODATE = Month(txtTodate.Text) & "/" & Microsoft.VisualBasic.Day(txtTodate.Text) & "/" & Year(txtTodate.Text)

                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Customer in ('" & _Customer & "') and Department in ('" & _Department & "') and Merchant in ('" & _Merchant & "') and left(r.Met_Des,8)='" & cboQuality.Text & "' and NC_Comment in ('15.BATCHES TB OVER DYED','16.STRIPPED TB OVER DYED','18.OFF SHADE BULK','19.OFF SHADE SAMPLE','20.OFF SHADE YARN DYE','21.WET FORM - TB REPROCESS')"
                        Else
                            vcWharer = " Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Customer in ('" & _Customer & "') and Department in ('" & _Department & "') and Merchant in ('" & _Merchant & "') and left(r.Met_Des,5)='" & cboQuality.Text & "' and NC_Comment in ('15.BATCHES TB OVER DYED','16.STRIPPED TB OVER DYED','18.OFF SHADE BULK','19.OFF SHADE SAMPLE','20.OFF SHADE YARN DYE','21.WET FORM - TB REPROCESS')"
                        End If
                    Else
                        vcWharer = " Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Customer in ('" & _Customer & "') and Department in ('" & _Department & "') and Merchant in ('" & _Merchant & "') and NC_Comment in ('15.BATCHES TB OVER DYED','16.STRIPPED TB OVER DYED','18.OFF SHADE BULK','19.OFF SHADE SAMPLE','20.OFF SHADE YARN DYE','21.WET FORM - TB REPROCESS') "

                    End If
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then

                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = "Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Customer in ('" & _Customer & "') and Department in ('" & _Department & "')  and left(r.Met_Des,8)='" & cboQuality.Text & "' and NC_Comment in ('15.BATCHES TB OVER DYED','16.STRIPPED TB OVER DYED','18.OFF SHADE BULK','19.OFF SHADE SAMPLE','20.OFF SHADE YARN DYE','21.WET FORM - TB REPROCESS')"
                        Else
                            vcWharer = "Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Customer in ('" & _Customer & "') and Department in ('" & _Department & "')  and left(r.Met_Des,5)='" & cboQuality.Text & "' and NC_Comment in ('15.BATCHES TB OVER DYED','16.STRIPPED TB OVER DYED','18.OFF SHADE BULK','19.OFF SHADE SAMPLE','20.OFF SHADE YARN DYE','21.WET FORM - TB REPROCESS') "
                        End If
                    Else
                        vcWharer = "Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Customer in ('" & _Customer & "') and Department in ('" & _Department & "') and NC_Comment in ('15.BATCHES TB OVER DYED','16.STRIPPED TB OVER DYED','18.OFF SHADE BULK','19.OFF SHADE SAMPLE','20.OFF SHADE YARN DYE','21.WET FORM - TB REPROCESS')"
                    End If
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = "Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Customer in ('" & _Customer & "')  and Merchant in ('" & _Merchant & "') and left(r.Met_Des,8)='" & cboQuality.Text & "' and NC_Comment in ('15.BATCHES TB OVER DYED','16.STRIPPED TB OVER DYED','18.OFF SHADE BULK','19.OFF SHADE SAMPLE','20.OFF SHADE YARN DYE','21.WET FORM - TB REPROCESS') "
                        Else
                            vcWharer = "Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Customer in ('" & _Customer & "')  and Merchant in ('" & _Merchant & "') and left(r.Met_Des,5)='" & cboQuality.Text & "' and NC_Comment in ('15.BATCHES TB OVER DYED','16.STRIPPED TB OVER DYED','18.OFF SHADE BULK','19.OFF SHADE SAMPLE','20.OFF SHADE YARN DYE','21.WET FORM - TB REPROCESS')"
                        End If
                    Else
                        vcWharer = "Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Customer in ('" & _Customer & "')  and Merchant in ('" & _Merchant & "') and NC_Comment in ('15.BATCHES TB OVER DYED','16.STRIPPED TB OVER DYED','18.OFF SHADE BULK','19.OFF SHADE SAMPLE','20.OFF SHADE YARN DYE','21.WET FORM - TB REPROCESS') "
                    End If
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then

                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Department in ('" & _Department & "') and Merchant in ('" & _Merchant & "') and left(r.Met_Des,8)='" & cboQuality.Text & "' and NC_Comment in ('15.BATCHES TB OVER DYED','16.STRIPPED TB OVER DYED','18.OFF SHADE BULK','19.OFF SHADE SAMPLE','20.OFF SHADE YARN DYE','21.WET FORM - TB REPROCESS')"
                        Else
                            vcWharer = " Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Department in ('" & _Department & "') and Merchant in ('" & _Merchant & "') and left(r.Met_Des,5)='" & cboQuality.Text & "' and NC_Comment in ('15.BATCHES TB OVER DYED','16.STRIPPED TB OVER DYED','18.OFF SHADE BULK','19.OFF SHADE SAMPLE','20.OFF SHADE YARN DYE','21.WET FORM - TB REPROCESS')"
                        End If
                    Else
                        vcWharer = " Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Department in ('" & _Department & "') and Merchant in ('" & _Merchant & "') and NC_Comment in ('15.BATCHES TB OVER DYED','16.STRIPPED TB OVER DYED','18.OFF SHADE BULK','19.OFF SHADE SAMPLE','20.OFF SHADE YARN DYE','21.WET FORM - TB REPROCESS') "
                    End If
                ElseIf Trim(txtCustomer.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Customer in ('" & _Customer & "')  and left(r.Met_Des,8)='" & cboQuality.Text & "' and NC_Comment in ('15.BATCHES TB OVER DYED','16.STRIPPED TB OVER DYED','18.OFF SHADE BULK','19.OFF SHADE SAMPLE','20.OFF SHADE YARN DYE','21.WET FORM - TB REPROCESS')"
                        Else
                            vcWharer = " Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Customer in ('" & _Customer & "')  and left(r.Met_Des,5)='" & cboQuality.Text & "' and NC_Comment in ('15.BATCHES TB OVER DYED','16.STRIPPED TB OVER DYED','18.OFF SHADE BULK','19.OFF SHADE SAMPLE','20.OFF SHADE YARN DYE','21.WET FORM - TB REPROCESS')"
                        End If
                    Else
                        vcWharer = " Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Customer in ('" & _Customer & "')and NC_Comment in ('15.BATCHES TB OVER DYED','16.STRIPPED TB OVER DYED','18.OFF SHADE BULK','19.OFF SHADE SAMPLE','20.OFF SHADE YARN DYE','21.WET FORM - TB REPROCESS') "

                    End If
                ElseIf Trim(txtDepartment.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and  Department in ('" & _Department & "')  left(r.Met_Des,8)='" & cboQuality.Text & "' and NC_Comment in ('15.BATCHES TB OVER DYED','16.STRIPPED TB OVER DYED','18.OFF SHADE BULK','19.OFF SHADE SAMPLE','20.OFF SHADE YARN DYE','21.WET FORM - TB REPROCESS')"
                        Else
                            vcWharer = " Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and  Department in ('" & _Department & "')  left(r.Met_Des,5)='" & cboQuality.Text & "' and NC_Comment in ('15.BATCHES TB OVER DYED','16.STRIPPED TB OVER DYED','18.OFF SHADE BULK','19.OFF SHADE SAMPLE','20.OFF SHADE YARN DYE','21.WET FORM - TB REPROCESS')"
                        End If
                    Else
                        vcWharer = " Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and  Department in ('" & _Department & "') and NC_Comment in ('15.BATCHES TB OVER DYED','16.STRIPPED TB OVER DYED','18.OFF SHADE BULK','19.OFF SHADE SAMPLE','20.OFF SHADE YARN DYE','21.WET FORM - TB REPROCESS') "
                    End If
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    If Trim(cboQuality.Text) <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = "Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'  and Merchant in ('" & _Merchant & "') and left(r.Met_Des,8)='" & cboQuality.Text & "' and NC_Comment in ('15.BATCHES TB OVER DYED','16.STRIPPED TB OVER DYED','18.OFF SHADE BULK','19.OFF SHADE SAMPLE','20.OFF SHADE YARN DYE','21.WET FORM - TB REPROCESS')"
                        Else
                            vcWharer = "Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'  and Merchant in ('" & _Merchant & "') and left(r.Met_Des,5)='" & cboQuality.Text & "' and NC_Comment in ('15.BATCHES TB OVER DYED','16.STRIPPED TB OVER DYED','18.OFF SHADE BULK','19.OFF SHADE SAMPLE','20.OFF SHADE YARN DYE','21.WET FORM - TB REPROCESS')"
                        End If
                    Else
                        vcWharer = "Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'  and Merchant in ('" & _Merchant & "') and NC_Comment in ('15.BATCHES TB OVER DYED','16.STRIPPED TB OVER DYED','18.OFF SHADE BULK','19.OFF SHADE SAMPLE','20.OFF SHADE YARN DYE','21.WET FORM - TB REPROCESS') "
                    End If
                Else
                    If Trim(cboQuality.Text) <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = "  Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and left(r.Met_Des,8)='" & cboQuality.Text & "' and NC_Comment in ('15.BATCHES TB OVER DYED','16.STRIPPED TB OVER DYED','18.OFF SHADE BULK','19.OFF SHADE SAMPLE','20.OFF SHADE YARN DYE','21.WET FORM - TB REPROCESS')"
                        Else
                            vcWharer = "  Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and left(r.Met_Des,5)='" & cboQuality.Text & "' and NC_Comment in ('15.BATCHES TB OVER DYED','16.STRIPPED TB OVER DYED','18.OFF SHADE BULK','19.OFF SHADE SAMPLE','20.OFF SHADE YARN DYE','21.WET FORM - TB REPROCESS') "
                        End If
                    Else
                        vcWharer = "  Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and NC_Comment in ('15.BATCHES TB OVER DYED','16.STRIPPED TB OVER DYED','18.OFF SHADE BULK','19.OFF SHADE SAMPLE','20.OFF SHADE YARN DYE','21.WET FORM - TB REPROCESS') "

                    End If
                End If
                I = 0
                T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetMEETING", New SqlParameter("@cQryType", "REC"), New SqlParameter("@vcWhereClause", vcWharer))
                'T01 = DBEngin.ExecuteDa
                _QtyKg = 0
                _QtyMtr = 0
                _Rowcount = 0
                For Each DTRow3 As DataRow In T01.Tables(0).Rows
                    Dim newRow As DataRow = c_dataCustomer1.NewRow
                    Dim _Tollaranz As Double

                    'For Each DTRow1 As DataRow In M01.Tables(0).Rows
                    If T01.Tables(0).Rows(I)("Del_Date") >= txtFromDate.Text Then

                        Dim _Material As String
                        _Material = T01.Tables(0).Rows(I)("Metrrial")
                        '_Material = Microsoft.VisualBasic.Right(_Material, 7)
                        '_Material = Microsoft.VisualBasic.Left(_Material, 2) & "-" & Microsoft.VisualBasic.Right(_Material, 5)
                        newRow("Sales Order") = T01.Tables(0).Rows(I)("Sales_Order")
                        newRow("Line Item") = T01.Tables(0).Rows(I)("Line_Item")
                        newRow("Material") = _Material
                        newRow("Description") = T01.Tables(0).Rows(I)("Met_Des")
                        newRow("Batch Qty (Mtr)") = Microsoft.VisualBasic.Format(T01.Tables(0).Rows(I)("PRD_Qty"), "#.00")
                        _QtyMtr = _QtyMtr + T01.Tables(0).Rows(I)("PRD_Qty")
                        newRow("Delivary Date") = Month(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Microsoft.VisualBasic.Day(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Year(T01.Tables(0).Rows(I)("Del_Date"))
                        _delivaryDT = Month(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Microsoft.VisualBasic.Day(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Year(T01.Tables(0).Rows(I)("Del_Date"))
                        newRow("Batch No") = T01.Tables(0).Rows(I)("Prduct_Order")
                        newRow("Customer") = T01.Tables(0).Rows(I)("Customer")

                        Sql = "select * from ZPP_DEL where Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and Line_Item='" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & "' and Delivery_Date='" & T01.Tables(0).Rows(I)("Del_Date") & "' and Product_Order='" & Trim(T01.Tables(0).Rows(I)("Prduct_Order")) & "'"
                        dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
                        If isValidDataset(dsUser) Then
                            newRow("Batch Qty (Kg)") = Microsoft.VisualBasic.Format(dsUser.Tables(0).Rows(0)("Order_Qty_Kg"), "#.00")
                            _QtyKg = _QtyKg + dsUser.Tables(0).Rows(0)("Order_Qty_Kg")
                            newRow("No.Of.Dys In S/L") = dsUser.Tables(0).Rows(0)("No_Day_Same_Opp")
                            newRow("NC Comment") = dsUser.Tables(0).Rows(0)("NC_Comment")
                            newRow("Next Operation") = dsUser.Tables(0).Rows(0)("Status")
                        End If

                        Sql = "select * from T07OTD_Comment1 where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07Material='" & _Material & "' and T07BatchNo='" & T01.Tables(0).Rows(I)("Prduct_Order") & "' order by T07Date DESC"
                        _OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                        If isValidDataset(_OTD) Then
                            If IsDate(_OTD.Tables(0).Rows(0)("T07Posible_Date")) Then
                                If Year(_OTD.Tables(0).Rows(0)("T07Posible_Date")) = "1900" Then
                                Else
                                    newRow("Possible Devivary Date") = _OTD.Tables(0).Rows(0)("T07Posible_Date")

                                End If
                            End If

                            newRow("Reason") = _OTD.Tables(0).Rows(0)("T07Reason")
                            If _OTD.Tables(0).Rows(0)("T07Date") = Today Then
                                newRow(_Todayupdate) = _OTD.Tables(0).Rows(0)("T07Comment")
                            Else
                                newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")
                            End If
                            newRow("LIB dep") = _OTD.Tables(0).Rows(0)("T07LIB")
                            _YesDate = Today.AddDays(-1)
                            Sql = "select * from T07OTD_Comment1 where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07BatchNo='" & T01.Tables(0).Rows(I)("Prduct_Order") & "' and T07Date<'" & Today & "' order by T07Date DESC"
                            _OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                            If isValidDataset(_OTD) Then
                                newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")
                            End If
                        Else
                            'Sql = "select * from T07OTD_Comment where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07Material='" & _Material & "'  order by T07Date DESC"
                            '_OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                            'If isValidDataset(_OTD) Then
                            '    '  newRow("Possible Devivary Date") = _OTD.Tables(0).Rows(0)("T07Posible_Date")
                            '    ' newRow("Reason") = _OTD.Tables(0).Rows(0)("T07Reason")
                            '    newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")
                            '    ' newRow("LIB dep") = _OTD.Tables(0).Rows(0)("T07LIB")
                            'End If
                        End If


                        newRow("Week") = "Week " & DatePart(DateInterval.WeekOfYear, _delivaryDT)
                        c_dataCustomer1.Rows.Add(newRow)
                        ' End If
                        _Rowcount = _Rowcount + 1
                        ' If Not x = M01.Tables(0).Rows.Count - 1 Then
                        ' rsT24RecHeader.MoveNext()
                    End If
                    I = I + 1
                Next
            ElseIf Trim(cboPO.Text) = "AW Preparation" Then
                Dim _FROMDATE As Date
                Dim _TODATE As Date
                Dim _OTD As DataSet

                _FROMDATE = Month(txtFromDate.Text) & "/" & Microsoft.VisualBasic.Day(txtFromDate.Text) & "/" & Year(txtFromDate.Text)
                _TODATE = Month(txtTodate.Text) & "/" & Microsoft.VisualBasic.Day(txtTodate.Text) & "/" & Year(txtTodate.Text)

                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw preparation') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and left(r.Met_Des,8)='" & cboQuality.Text & "' "
                        Else
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw preparation') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw preparation') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')"

                    End If
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw preparation') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw preparation') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')  and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw preparation') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') "

                    End If

                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw preparation') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Merchant in ('" & _Merchant & "') and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw preparation') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "') and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw preparation') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "')"

                    End If
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then

                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw preparation') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and  r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw preparation') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and  r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw preparation') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and  r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')"

                    End If

                ElseIf Trim(txtCustomer.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw preparation') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "')  and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw preparation') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "')  and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw preparation') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') "

                    End If

                ElseIf Trim(txtDepartment.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw preparation') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Department in ('" & _Department & "')  and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw preparation') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Department in ('" & _Department & "')  and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw preparation') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and  r.Department in ('" & _Department & "') "

                    End If
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw preparation') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and r.Merchant in ('" & _Merchant & "') and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw preparation') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Merchant in ('" & _Merchant & "') and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw preparation') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and r.Merchant in ('" & _Merchant & "')"

                    End If
                Else
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw preparation') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw preparation') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw preparation') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') "

                    End If
                End If
                I = 0
                'T01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                _QtyKg = 0
                _QtyMtr = 0
                _Rowcount = 0
                T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetMEETING", New SqlParameter("@cQryType", "TBP"), New SqlParameter("@vcWhereClause", vcWharer))
                'T01 = DBEngin.ExecuteDa
                For Each DTRow3 As DataRow In T01.Tables(0).Rows
                    Dim newRow As DataRow = c_dataCustomer1.NewRow
                    Dim _Tollaranz As Double

                    'For Each DTRow1 As DataRow In M01.Tables(0).Rows


                    Dim _Material As String
                    _Material = T01.Tables(0).Rows(I)("Metrrial")
                    '_Material = Microsoft.VisualBasic.Right(_Material, 7)
                    '_Material = Microsoft.VisualBasic.Left(_Material, 2) & "-" & Microsoft.VisualBasic.Right(_Material, 5)
                    newRow("Sales Order") = T01.Tables(0).Rows(I)("Sales_Order")
                    newRow("Line Item") = T01.Tables(0).Rows(I)("Line_Item")
                    newRow("Material") = _Material
                    newRow("Description") = T01.Tables(0).Rows(I)("Met_Des")
                    newRow("Batch Qty (Mtr)") = Microsoft.VisualBasic.Format(T01.Tables(0).Rows(I)("PRD_Qty"), "#.00")
                    _QtyMtr = _QtyMtr + T01.Tables(0).Rows(I)("PRD_Qty")
                    newRow("Delivary Date") = Month(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Microsoft.VisualBasic.Day(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Year(T01.Tables(0).Rows(I)("Del_Date"))
                    _delivaryDT = Month(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Microsoft.VisualBasic.Day(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Year(T01.Tables(0).Rows(I)("Del_Date"))
                    newRow("Batch No") = T01.Tables(0).Rows(I)("Prduct_Order")
                    newRow("Customer") = T01.Tables(0).Rows(I)("Customer")

                    Sql = "select * from ZPP_DEL where Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and Line_Item='" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & "' and Delivery_Date='" & T01.Tables(0).Rows(I)("Del_Date") & "' and Product_Order='" & Trim(T01.Tables(0).Rows(I)("Prduct_Order")) & "'"
                    dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(dsUser) Then
                        newRow("Batch Qty (Kg)") = Microsoft.VisualBasic.Format(dsUser.Tables(0).Rows(0)("Order_Qty_Kg"), "#.00")
                        _QtyKg = _QtyKg + dsUser.Tables(0).Rows(0)("Order_Qty_Kg")
                        newRow("No.Of.Dys In S/L") = dsUser.Tables(0).Rows(0)("No_Day_Same_Opp")
                        newRow("NC Comment") = dsUser.Tables(0).Rows(0)("NC_Comment")
                        newRow("Next Operation") = dsUser.Tables(0).Rows(0)("Status")
                    End If

                    Sql = "select * from T07OTD_Comment1 where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07Material='" & _Material & "' and T07BatchNo='" & T01.Tables(0).Rows(I)("Prduct_Order") & "' order by T07Date DESC"
                    _OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(_OTD) Then
                        If IsDate(_OTD.Tables(0).Rows(0)("T07Posible_Date")) Then
                            If Year(_OTD.Tables(0).Rows(0)("T07Posible_Date")) = "1900" Then
                            Else
                                newRow("Possible Devivary Date") = _OTD.Tables(0).Rows(0)("T07Posible_Date")

                            End If
                        End If


                        newRow("Reason") = _OTD.Tables(0).Rows(0)("T07Reason")
                        If _OTD.Tables(0).Rows(0)("T07date") = Today Then
                            newRow(_Todayupdate) = _OTD.Tables(0).Rows(0)("T07Comment")
                        Else
                            newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")

                            Dim _DyeMC1 As String

                            Sql = "select * from FR_Update where Batch_No='" & Trim(T01.Tables(0).Rows(I)("Prduct_Order")) & "'"
                            _OTDComm = DBEngin.ExecuteDataset(con, Nothing, Sql)
                            If isValidDataset(_OTDComm) Then
                                If Not DBNull.Value.Equals(_OTDComm.Tables(0).Rows(0)("Dye_Machine")) Then
                                    _DyeMC1 = _OTDComm.Tables(0).Rows(0)("Dye_Machine")
                                    newRow(_Todayupdate) = "Dye Pln on - " & Month(_OTDComm.Tables(0).Rows(0)("Dye_Pln_Date")) & "/" & Microsoft.VisualBasic.Day(_OTDComm.Tables(0).Rows(0)("Dye_Pln_Date")) & "/" & Year(_OTDComm.Tables(0).Rows(0)("Dye_Pln_Date")) & "-" & _OTDComm.Tables(0).Rows(0)("Dye_Machine")
                                Else
                                    newRow(_Todayupdate) = "Dye pln on - " & Month(_OTDComm.Tables(0).Rows(0)("Dye_Pln_Date")) & "/" & Microsoft.VisualBasic.Day(_OTDComm.Tables(0).Rows(0)("Dye_Pln_Date")) & "/" & Year(_OTDComm.Tables(0).Rows(0)("Dye_Pln_Date"))
                                End If
                            Else
                                newRow(_Todayupdate) = "Pln not yet "
                            End If

                        End If
                        newRow("LIB dep") = _OTD.Tables(0).Rows(0)("T07LIB")

                        _YesDate = Today.AddDays(-1)
                        Sql = "select * from T07OTD_Comment1 where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07BatchNo='" & T01.Tables(0).Rows(I)("Prduct_Order") & "' and T07Date<'" & Today & "' order by T07Date DESC"
                        _OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                        If isValidDataset(_OTD) Then
                            newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")
                        End If

                    Else
                        'Sql = "select * from T07OTD_Comment where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07Material='" & _Material & "'  order by T07Date DESC"
                        '_OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                        'If isValidDataset(_OTD) Then
                        '    'newRow("Possible Devivary Date") = _OTD.Tables(0).Rows(0)("T07Posible_Date")
                        '    'newRow("Reason") = _OTD.Tables(0).Rows(0)("T07Reason")
                        '    newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")
                        '    'newRow("LIB dep") = _OTD.Tables(0).Rows(0)("T07LIB")
                        'End If
                        Dim _DyeMC1 As String

                        Sql = "select * from FR_Update where Batch_No='" & Trim(T01.Tables(0).Rows(I)("Prduct_Order")) & "'"
                        _OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                        If isValidDataset(_OTD) Then
                            If Not DBNull.Value.Equals(_OTD.Tables(0).Rows(0)("Dye_Machine")) Then
                                _DyeMC1 = _OTD.Tables(0).Rows(0)("Dye_Machine")
                                newRow(_Todayupdate) = "Dye Pln on - " & Month(_OTD.Tables(0).Rows(0)("Dye_Pln_Date")) & "/" & Microsoft.VisualBasic.Day(_OTD.Tables(0).Rows(0)("Dye_Pln_Date")) & "/" & Year(_OTD.Tables(0).Rows(0)("Dye_Pln_Date")) & "-" & _OTD.Tables(0).Rows(0)("Dye_Machine")
                            Else
                                newRow(_Todayupdate) = "Dye pln on - " & Month(_OTD.Tables(0).Rows(0)("Dye_Pln_Date")) & "/" & Microsoft.VisualBasic.Day(_OTD.Tables(0).Rows(0)("Dye_Pln_Date")) & "/" & Year(_OTD.Tables(0).Rows(0)("Dye_Pln_Date"))
                            End If
                        Else
                            newRow(_Todayupdate) = "Pln not yet "
                        End If
                    End If

                    newRow("Week") = "Week " & DatePart(DateInterval.WeekOfYear, _delivaryDT)
                    c_dataCustomer1.Rows.Add(newRow)
                    ' End If

                    ' If Not x = M01.Tables(0).Rows.Count - 1 Then
                    ' rsT24RecHeader.MoveNext()
                    _Rowcount = _Rowcount + 1
                    I = I + 1
                Next
            ElseIf Trim(cboPO.Text) = "Aw Presetting" Then
                Dim _FROMDATE As Date
                Dim _TODATE As Date
                Dim _OTD As DataSet

                _FROMDATE = Month(txtFromDate.Text) & "/" & Microsoft.VisualBasic.Day(txtFromDate.Text) & "/" & Year(txtFromDate.Text)
                _TODATE = Month(txtTodate.Text) & "/" & Microsoft.VisualBasic.Day(txtTodate.Text) & "/" & Year(txtTodate.Text)

                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw Presetting') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw Presetting') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw Presetting') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')"

                    End If
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw Presetting') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')   and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw Presetting') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')   and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw Presetting') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') "

                    End If
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw Presetting') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Merchant in ('" & _Merchant & "')  and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw Presetting') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "')  and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw Presetting') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "')"

                    End If
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then

                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw Presetting') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw Presetting') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw Presetting') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')"

                    End If

                ElseIf Trim(txtCustomer.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw Presetting') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "')   and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw Presetting') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "')  and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw Presetting') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') "

                    End If

                ElseIf Trim(txtDepartment.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw Presetting') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Department in ('" & _Department & "')   and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw Presetting') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Department in ('" & _Department & "')   and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw Presetting') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Department in ('" & _Department & "')"

                    End If
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw Presetting') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Merchant in ('" & _Merchant & "')  and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw Presetting') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Merchant in ('" & _Merchant & "')  and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw Presetting') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Merchant in ('" & _Merchant & "')"

                    End If
                Else
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw Presetting') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw Presetting') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw Presetting') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') "

                    End If
                End If
                I = 0
                'T01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetMEETING", New SqlParameter("@cQryType", "TBP"), New SqlParameter("@vcWhereClause", vcWharer))
                'T01 = DBEngin.ExecuteDa
                _QtyKg = 0
                _QtyMtr = 0
                _Rowcount = 0
                For Each DTRow3 As DataRow In T01.Tables(0).Rows
                    Dim newRow As DataRow = c_dataCustomer1.NewRow
                    Dim _Tollaranz As Double

                    'For Each DTRow1 As DataRow In M01.Tables(0).Rows


                    Dim _Material As String
                    _Material = T01.Tables(0).Rows(I)("Metrrial")
                    '_Material = Microsoft.VisualBasic.Right(_Material, 7)
                    '_Material = Microsoft.VisualBasic.Left(_Material, 2) & "-" & Microsoft.VisualBasic.Right(_Material, 5)
                    newRow("Sales Order") = T01.Tables(0).Rows(I)("Sales_Order")
                    newRow("Line Item") = T01.Tables(0).Rows(I)("Line_Item")
                    newRow("Material") = _Material
                    newRow("Description") = T01.Tables(0).Rows(I)("Met_Des")
                    newRow("Batch Qty (Mtr)") = Microsoft.VisualBasic.Format(T01.Tables(0).Rows(I)("PRD_Qty"), "#.00")
                    _QtyMtr = _QtyMtr + T01.Tables(0).Rows(I)("PRD_Qty")
                    newRow("Delivary Date") = Month(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Microsoft.VisualBasic.Day(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Year(T01.Tables(0).Rows(I)("Del_Date"))

                    _delivaryDT = Month(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Microsoft.VisualBasic.Day(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Year(T01.Tables(0).Rows(I)("Del_Date"))
                    newRow("Batch No") = T01.Tables(0).Rows(I)("Prduct_Order")
                    newRow("Customer") = T01.Tables(0).Rows(I)("Customer")

                    Sql = "select * from ZPP_DEL where Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and Line_Item='" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & "' and Delivery_Date='" & T01.Tables(0).Rows(I)("Del_Date") & "' and Product_Order='" & Trim(T01.Tables(0).Rows(I)("Prduct_Order")) & "'"
                    dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(dsUser) Then
                        newRow("Batch Qty (Kg)") = Microsoft.VisualBasic.Format(dsUser.Tables(0).Rows(0)("Order_Qty_Kg"), "#.00")
                        _QtyKg = _QtyKg + dsUser.Tables(0).Rows(0)("Order_Qty_Kg")
                        newRow("No.Of.Dys In S/L") = dsUser.Tables(0).Rows(0)("No_Day_Same_Opp")
                        newRow("NC Comment") = dsUser.Tables(0).Rows(0)("NC_Comment")
                        newRow("Next Operation") = dsUser.Tables(0).Rows(0)("Status")
                    End If

                    Sql = "select * from T07OTD_Comment1 where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07Material='" & _Material & "' and T07BatchNo='" & T01.Tables(0).Rows(I)("Prduct_Order") & "' order by T07Date DESC"
                    _OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(_OTD) Then
                        If IsDate(_OTD.Tables(0).Rows(0)("T07Posible_Date")) Then
                            If Year(_OTD.Tables(0).Rows(0)("T07Posible_Date")) = "1900" Then
                            Else
                                newRow("Possible Devivary Date") = _OTD.Tables(0).Rows(0)("T07Posible_Date")

                            End If
                        End If

                        newRow("Reason") = _OTD.Tables(0).Rows(0)("T07Reason")
                        If _OTD.Tables(0).Rows(0)("T07date") = Today Then
                            newRow(_Todayupdate) = _OTD.Tables(0).Rows(0)("T07Comment")

                            Dim _DyeMC1 As String

                            Sql = "select * from FR_Update where Batch_No='" & Trim(T01.Tables(0).Rows(I)("Prduct_Order")) & "'"
                            _OTDComm = DBEngin.ExecuteDataset(con, Nothing, Sql)
                            If isValidDataset(_OTDComm) Then
                                If Not DBNull.Value.Equals(_OTDComm.Tables(0).Rows(0)("Dye_Machine")) Then
                                    _DyeMC1 = _OTDComm.Tables(0).Rows(0)("Dye_Machine")
                                    newRow(_Todayupdate) = "Dye Pln on - " & Month(_OTDComm.Tables(0).Rows(0)("Dye_Pln_Date")) & "/" & Microsoft.VisualBasic.Day(_OTDComm.Tables(0).Rows(0)("Dye_Pln_Date")) & "/" & Year(_OTDComm.Tables(0).Rows(0)("Dye_Pln_Date")) & "-" & _OTDComm.Tables(0).Rows(0)("Dye_Machine")
                                Else
                                    newRow(_Todayupdate) = "Dye pln on - " & Month(_OTDComm.Tables(0).Rows(0)("Dye_Pln_Date")) & "/" & Microsoft.VisualBasic.Day(_OTDComm.Tables(0).Rows(0)("Dye_Pln_Date")) & "/" & Year(_OTDComm.Tables(0).Rows(0)("Dye_Pln_Date"))
                                End If
                            Else
                                newRow(_Todayupdate) = "Pln not yet "
                            End If

                        Else
                            newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")
                        End If
                        newRow("LIB dep") = _OTD.Tables(0).Rows(0)("T07LIB")

                        _YesDate = Today.AddDays(-1)
                        Sql = "select * from T07OTD_Comment1 where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07BatchNo='" & T01.Tables(0).Rows(I)("Prduct_Order") & "' and T07Date<'" & Today & "' order by T07Date DESC"
                        _OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                        If isValidDataset(_OTD) Then
                            newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")
                        End If

                    Else
                        Dim _DyeMC1 As String

                        Sql = "select * from FR_Update where Batch_No='" & Trim(T01.Tables(0).Rows(I)("Prduct_Order")) & "'"
                        _OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                        If isValidDataset(_OTD) Then
                            If Not DBNull.Value.Equals(_OTD.Tables(0).Rows(0)("Dye_Machine")) Then
                                _DyeMC1 = _OTD.Tables(0).Rows(0)("Dye_Machine")
                                newRow(_Todayupdate) = "Dye pln on - " & Month(_OTD.Tables(0).Rows(0)("Dye_Pln_Date")) & "/" & Microsoft.VisualBasic.Day(_OTD.Tables(0).Rows(0)("Dye_Pln_Date")) & "/" & Year(_OTD.Tables(0).Rows(0)("Dye_Pln_Date")) & "-" & _OTD.Tables(0).Rows(0)("Dye_Machine")
                            Else
                                newRow(_Todayupdate) = "Dye pln on - " & Month(_OTD.Tables(0).Rows(0)("Dye_Pln_Date")) & "/" & Microsoft.VisualBasic.Day(_OTD.Tables(0).Rows(0)("Dye_Pln_Date")) & "/" & Year(_OTD.Tables(0).Rows(0)("Dye_Pln_Date"))
                            End If
                        Else
                            newRow(_Todayupdate) = "Pln not yet"
                        End If
                    End If

                    newRow("Week") = "Week " & DatePart(DateInterval.WeekOfYear, _delivaryDT)
                    c_dataCustomer1.Rows.Add(newRow)
                    ' End If

                    ' If Not x = M01.Tables(0).Rows.Count - 1 Then
                    ' rsT24RecHeader.MoveNext()
                    _Rowcount = _Rowcount + 1
                    I = I + 1
                Next
            ElseIf Trim(cboPO.Text) = "Aw Dyeing" Then
                Dim _FROMDATE As Date
                Dim _TODATE As Date
                Dim _OTD As DataSet

                _FROMDATE = Month(txtFromDate.Text) & "/" & Microsoft.VisualBasic.Day(txtFromDate.Text) & "/" & Year(txtFromDate.Text)
                _TODATE = Month(txtTodate.Text) & "/" & Microsoft.VisualBasic.Day(txtTodate.Text) & "/" & Year(txtTodate.Text)

                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Dye') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.RESCHEDULE – WASH','24.RESCHEDULE – STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RE SCHEDULE – SAMPLE','')  and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Dye') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.RESCHEDULE – WASH','24.RESCHEDULE – STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RE SCHEDULE – SAMPLE','')  and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Dye') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.RESCHEDULE – WASH','24.RESCHEDULE – STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RE SCHEDULE – SAMPLE','') "

                    End If
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Dye') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')  r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.RESCHEDULE – WASH','24.RESCHEDULE – STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RE SCHEDULE – SAMPLE','')  and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Dye') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') ANd r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.RESCHEDULE – WASH','24.RESCHEDULE – STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RE SCHEDULE – SAMPLE','') and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Dye') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')  and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.RESCHEDULE – WASH','24.RESCHEDULE – STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RE SCHEDULE – SAMPLE','')"

                    End If
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Dye') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "') and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.RESCHEDULE – WASH','24.RESCHEDULE – STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RE SCHEDULE – SAMPLE','') and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Dye') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "') and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.RESCHEDULE – WASH','24.RESCHEDULE – STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RE SCHEDULE – SAMPLE','')  and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Dye') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "') and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.RESCHEDULE – WASH','24.RESCHEDULE – STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RE SCHEDULE – SAMPLE','') "

                    End If
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then

                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Dye') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.RESCHEDULE – WASH','24.RESCHEDULE – STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RE SCHEDULE – SAMPLE','') and  left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Dye') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.RESCHEDULE – WASH','24.RESCHEDULE – STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RE SCHEDULE – SAMPLE','') and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Dye') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.RESCHEDULE – WASH','24.RESCHEDULE – STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RE SCHEDULE – SAMPLE','')"

                    End If

                ElseIf Trim(txtCustomer.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Dye') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "')  and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.RESCHEDULE – WASH','24.RESCHEDULE – STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RE SCHEDULE – SAMPLE','')  and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Dye') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "')  and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.RESCHEDULE – WASH','24.RESCHEDULE – STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RE SCHEDULE – SAMPLE','')  and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Dye') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "')  and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.RESCHEDULE – WASH','24.RESCHEDULE – STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RE SCHEDULE – SAMPLE','') "

                    End If

                ElseIf Trim(txtDepartment.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Dye') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and  r.Department in ('" & _Department & "')  and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.RESCHEDULE – WASH','24.RESCHEDULE – STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RE SCHEDULE – SAMPLE','')  and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Dye') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and  r.Department in ('" & _Department & "')  and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.RESCHEDULE – WASH','24.RESCHEDULE – STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RE SCHEDULE – SAMPLE','')  and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Dye') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Department in ('" & _Department & "') and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.RESCHEDULE – WASH','24.RESCHEDULE – STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RE SCHEDULE – SAMPLE','') "

                    End If
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Dye') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Merchant in ('" & _Merchant & "') and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.RESCHEDULE – WASH','24.RESCHEDULE – STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RE SCHEDULE – SAMPLE','')  and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Dye') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Merchant in ('" & _Merchant & "') and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.RESCHEDULE – WASH','24.RESCHEDULE – STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RE SCHEDULE – SAMPLE','') and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Dye') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Merchant in ('" & _Merchant & "') and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.RESCHEDULE – WASH','24.RESCHEDULE – STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RE SCHEDULE – SAMPLE','') "

                    End If
                Else
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Dye') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.RESCHEDULE – WASH','24.RESCHEDULE – STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RE SCHEDULE – SAMPLE','') and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Dye') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Merchant in ('" & _Merchant & "') and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.RESCHEDULE – WASH','24.RESCHEDULE – STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RE SCHEDULE – SAMPLE','')  and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Dye') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.RESCHEDULE – WASH','24.RESCHEDULE – STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RE SCHEDULE – SAMPLE','') "

                    End If
                End If
                I = 0
                'T01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetMEETING", New SqlParameter("@cQryType", "DYE"), New SqlParameter("@vcWhereClause", vcWharer))
                'T01 = DBEngin.ExecuteDa
                _Rowcount = 0
                _QtyKg = 0
                _QtyMtr = 0
                Dim _DyeMC1 As String
                _DyeMC1 = ""
                For Each DTRow3 As DataRow In T01.Tables(0).Rows
                    Dim newRow As DataRow = c_dataCustomer1.NewRow
                    Dim _Tollaranz As Double

                    'For Each DTRow1 As DataRow In M01.Tables(0).Rows
                    Sql = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,f.Stock_Code,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date inner join FR_Update F on f.Batch_No=r.Prduct_Order where s.Status='Fales' and r.Prduct_Order='" & Trim(T01.Tables(0).Rows(I)("Prduct_Order")) & "' and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') "
                    T03 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(T03) Then
                    Else
                        Dim _Material As String
                        _Material = T01.Tables(0).Rows(I)("Metrrial")
                        '_Material = Microsoft.VisualBasic.Right(_Material, 7)
                        '_Material = Microsoft.VisualBasic.Left(_Material, 2) & "-" & Microsoft.VisualBasic.Right(_Material, 5)
                        newRow("Sales Order") = T01.Tables(0).Rows(I)("Sales_Order")
                        newRow("Line Item") = T01.Tables(0).Rows(I)("Line_Item")
                        newRow("Material") = _Material
                        newRow("Description") = T01.Tables(0).Rows(I)("Met_Des")
                        newRow("Batch Qty (Mtr)") = Microsoft.VisualBasic.Format(T01.Tables(0).Rows(I)("PRD_Qty"), "#.00")
                        _QtyMtr = _QtyMtr + T01.Tables(0).Rows(I)("PRD_Qty")
                        newRow("Delivary Date") = Month(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Microsoft.VisualBasic.Day(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Year(T01.Tables(0).Rows(I)("Del_Date"))
                        _delivaryDT = Month(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Microsoft.VisualBasic.Day(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Year(T01.Tables(0).Rows(I)("Del_Date"))
                        newRow("Batch No") = T01.Tables(0).Rows(I)("Prduct_Order")
                        newRow("Customer") = T01.Tables(0).Rows(I)("Customer")

                        Sql = "select * from ZPP_DEL where Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and Line_Item='" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & "' and Delivery_Date='" & T01.Tables(0).Rows(I)("Del_Date") & "' and Product_Order='" & Trim(T01.Tables(0).Rows(I)("Prduct_Order")) & "'"
                        dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
                        If isValidDataset(dsUser) Then
                            newRow("Batch Qty (Kg)") = Microsoft.VisualBasic.Format(dsUser.Tables(0).Rows(0)("Order_Qty_Kg"), "#.00")
                            _QtyKg = _QtyKg + dsUser.Tables(0).Rows(0)("Order_Qty_Kg")
                            newRow("No.Of.Dys In S/L") = dsUser.Tables(0).Rows(0)("No_Day_Same_Opp")
                            newRow("NC Comment") = dsUser.Tables(0).Rows(0)("NC_Comment")
                            newRow("Next Operation") = dsUser.Tables(0).Rows(0)("Status")
                        End If

                        _DyeMC1 = ""
                        Sql = "select * from T07OTD_Comment1 where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07Material='" & _Material & "' and T07BatchNo='" & T01.Tables(0).Rows(I)("Prduct_Order") & "' order by T07Date DESC"
                        _OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                        If isValidDataset(_OTD) Then
                            If IsDate(_OTD.Tables(0).Rows(0)("T07Posible_Date")) Then
                                If Year(_OTD.Tables(0).Rows(0)("T07Posible_Date")) = "1900" Then
                                Else
                                    newRow("Possible Devivary Date") = _OTD.Tables(0).Rows(0)("T07Posible_Date")

                                End If
                            End If

                            newRow("Reason") = _OTD.Tables(0).Rows(0)("T07Reason")
                            newRow("LIB dep") = _OTD.Tables(0).Rows(0)("T07LIB")
                            If _OTD.Tables(0).Rows(0)("T07date") = Today Then
                                newRow(_Todayupdate) = _OTD.Tables(0).Rows(0)("T07Comment")
                            Else
                                newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")
                            End If

                            Sql = "select * from FR_Update where Batch_No='" & Trim(T01.Tables(0).Rows(I)("Prduct_Order")) & "'"
                            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
                            If isValidDataset(dsUser) Then
                                If Not DBNull.Value.Equals(dsUser.Tables(0).Rows(0)("Dye_Machine")) Then
                                    _DyeMC1 = dsUser.Tables(0).Rows(0)("Dye_Machine")
                                    If _OTD.Tables(0).Rows(0)("T07date") = Today Then
                                        newRow(_Todayupdate) = _OTD.Tables(0).Rows(0)("T07Comment")
                                    Else
                                        newRow(_Todayupdate) = "Dye Plan on " & Month(dsUser.Tables(0).Rows(0)("Dye_Pln_Date")) & "/" & Microsoft.VisualBasic.Day(dsUser.Tables(0).Rows(0)("Dye_Pln_Date")) & "/" & Year(dsUser.Tables(0).Rows(0)("Dye_Pln_Date")) & "-" & dsUser.Tables(0).Rows(0)("Dye_Machine")
                                    End If
                                Else
                                    If _OTD.Tables(0).Rows(0)("T07date") = Today Then
                                        newRow(_Todayupdate) = _OTD.Tables(0).Rows(0)("T07Comment")
                                    Else
                                        newRow(_Todayupdate) = "Dye Plan on " & Month(dsUser.Tables(0).Rows(0)("Dye_Pln_Date")) & "/" & Microsoft.VisualBasic.Day(dsUser.Tables(0).Rows(0)("Dye_Pln_Date")) & "/" & Year(dsUser.Tables(0).Rows(0)("Dye_Pln_Date"))
                                    End If
                                End If
                            End If

                            _YesDate = Today.AddDays(-1)
                            Sql = "select * from T07OTD_Comment1 where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07BatchNo='" & T01.Tables(0).Rows(I)("Prduct_Order") & "' and T07Date<'" & Today & "' order by T07Date DESC"
                            _OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                            If isValidDataset(_OTD) Then
                                newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")
                            End If

                        Else
                            Sql = "select * from T07OTD_Comment1 where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07Material='" & _Material & "'  order by T07Date DESC"
                            _OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                            If isValidDataset(_OTD) Then
                                'newRow("Possible Devivary Date") = _OTD.Tables(0).Rows(0)("T07Posible_Date")
                                'newRow("Reason") = _OTD.Tables(0).Rows(0)("T07Reason")
                                If _OTD.Tables(0).Rows(0)("T07date") = Today Then
                                    newRow(_Todayupdate) = _OTD.Tables(0).Rows(0)("T07Comment")
                                Else
                                    newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")
                                End If

                                _YesDate = Today.AddDays(-1)
                                Sql = "select * from T07OTD_Comment1 where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07BatchNo='" & T01.Tables(0).Rows(I)("Prduct_Order") & "' and T07Date<'" & Today & "' order by T07Date DESC"
                                _OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                                If isValidDataset(_OTD) Then
                                    newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")
                                End If

                            Else
                                Sql = "select * from FR_Update where Batch_No='" & Trim(T01.Tables(0).Rows(I)("Prduct_Order")) & "'"
                                _OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                                If isValidDataset(_OTD) Then
                                    If Not DBNull.Value.Equals(_OTD.Tables(0).Rows(0)("Dye_Machine")) Then
                                        _DyeMC1 = _OTD.Tables(0).Rows(0)("Dye_Machine")
                                        newRow(_Todayupdate) = "Dye Plan on " & Month(_OTD.Tables(0).Rows(0)("Dye_Pln_Date")) & "/" & Microsoft.VisualBasic.Day(_OTD.Tables(0).Rows(0)("Dye_Pln_Date")) & "/" & Year(_OTD.Tables(0).Rows(0)("Dye_Pln_Date")) & "-" & _OTD.Tables(0).Rows(0)("Dye_Machine")
                                    Else
                                        newRow(_Todayupdate) = "Dye Plan on " & Month(_OTD.Tables(0).Rows(0)("Dye_Pln_Date")) & "/" & Microsoft.VisualBasic.Day(_OTD.Tables(0).Rows(0)("Dye_Pln_Date")) & "/" & Year(_OTD.Tables(0).Rows(0)("Dye_Pln_Date"))
                                    End If
                                End If
                                'newRow("LIB dep") = _OTD.Tables(0).Rows(0)("T07LIB")
                            End If
                        End If

                        newRow("Week") = "Week " & DatePart(DateInterval.WeekOfYear, _delivaryDT)
                        c_dataCustomer1.Rows.Add(newRow)
                    End If
                    Dim _Rcount As Integer
                    _Rcount = UltraGrid4.Rows.Count
                    If Microsoft.VisualBasic.Left(_DyeMC1, 1) = "S" Or Microsoft.VisualBasic.Left(_DyeMC1, 1) = "Y" Or Microsoft.VisualBasic.Left(_DyeMC1, 1) = "W" Or IsNumeric(_DyeMC1) Then
                        UltraGrid4.Rows(_Rcount - 1).Cells(0).Appearance.BackColor = Color.Yellow
                        UltraGrid4.Rows(_Rcount - 1).Cells(1).Appearance.BackColor = Color.Yellow
                        UltraGrid4.Rows(_Rcount - 1).Cells(2).Appearance.BackColor = Color.Yellow
                        UltraGrid4.Rows(_Rcount - 1).Cells(3).Appearance.BackColor = Color.Yellow
                        UltraGrid4.Rows(_Rcount - 1).Cells(4).Appearance.BackColor = Color.Yellow
                        UltraGrid4.Rows(_Rcount - 1).Cells(5).Appearance.BackColor = Color.Yellow
                        UltraGrid4.Rows(_Rcount - 1).Cells(6).Appearance.BackColor = Color.Yellow
                        UltraGrid4.Rows(_Rcount - 1).Cells(7).Appearance.BackColor = Color.Yellow
                        UltraGrid4.Rows(_Rcount - 1).Cells(8).Appearance.BackColor = Color.Yellow
                        UltraGrid4.Rows(_Rcount - 1).Cells(9).Appearance.BackColor = Color.Yellow
                        UltraGrid4.Rows(_Rcount - 1).Cells(10).Appearance.BackColor = Color.Yellow
                        UltraGrid4.Rows(_Rcount - 1).Cells(11).Appearance.BackColor = Color.Yellow
                        UltraGrid4.Rows(_Rcount - 1).Cells(12).Appearance.BackColor = Color.Yellow
                        UltraGrid4.Rows(_Rcount - 1).Cells(13).Appearance.BackColor = Color.Yellow
                        UltraGrid4.Rows(_Rcount - 1).Cells(14).Appearance.BackColor = Color.Yellow
                        UltraGrid4.Rows(_Rcount - 1).Cells(15).Appearance.BackColor = Color.Yellow
                        UltraGrid4.Rows(_Rcount - 1).Cells(16).Appearance.BackColor = Color.Yellow

                    End If
                    ' If Not x = M01.Tables(0).Rows.Count - 1 Then
                    ' rsT24RecHeader.MoveNext()
                    _Rowcount = _Rowcount + 1
                    I = I + 1
                Next


                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Finishing') and   s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.RESCHEDULE – WASH','24.RESCHEDULE – STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RE SCHEDULE – SAMPLE')  and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Finishing') and   s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.RESCHEDULE – WASH','24.RESCHEDULE – STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RE SCHEDULE – SAMPLE')  and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Finishing') and   s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.RESCHEDULE – WASH','24.RESCHEDULE – STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RE SCHEDULE – SAMPLE')  "

                    End If
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Finishing') and   s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')  and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.RESCHEDULE – WASH','24.RESCHEDULE – STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RE SCHEDULE – SAMPLE')  and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Finishing') and   s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.RESCHEDULE – WASH','24.RESCHEDULE – STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RE SCHEDULE – SAMPLE')  and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Finishing') and   s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')  and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.RESCHEDULE – WASH','24.RESCHEDULE – STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RE SCHEDULE – SAMPLE') "

                    End If
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Finishing') and   s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "') and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.RESCHEDULE – WASH','24.RESCHEDULE – STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RE SCHEDULE – SAMPLE')  and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Finishing') and   s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "') and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.RESCHEDULE – WASH','24.RESCHEDULE – STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RE SCHEDULE – SAMPLE')  and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Finishing') and   s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "') and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.RESCHEDULE – WASH','24.RESCHEDULE – STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RE SCHEDULE – SAMPLE')  "

                    End If
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then

                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Finishing') and    r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.RESCHEDULE – WASH','24.RESCHEDULE – STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RE SCHEDULE – SAMPLE')  and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Finishing') and   r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.RESCHEDULE – WASH','24.RESCHEDULE – STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RE SCHEDULE – SAMPLE')  and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Finishing') and   r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.RESCHEDULE – WASH','24.RESCHEDULE – STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RE SCHEDULE – SAMPLE')  "

                    End If

                ElseIf Trim(txtCustomer.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Finishing') and   s.Customer in ('" & _Customer & "')  and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.RESCHEDULE – WASH','24.RESCHEDULE – STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RE SCHEDULE – SAMPLE')  and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Finishing') and   s.Customer in ('" & _Customer & "')  and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.RESCHEDULE – WASH','24.RESCHEDULE – STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RE SCHEDULE – SAMPLE')  and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Finishing') and   s.Customer in ('" & _Customer & "') and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.RESCHEDULE – WASH','24.RESCHEDULE – STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RE SCHEDULE – SAMPLE')  "

                    End If

                ElseIf Trim(txtDepartment.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Finishing') and    r.Department in ('" & _Department & "')  and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.RESCHEDULE – WASH','24.RESCHEDULE – STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RE SCHEDULE – SAMPLE')  and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Finishing') and    r.Department in ('" & _Department & "')  and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.RESCHEDULE – WASH','24.RESCHEDULE – STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RE SCHEDULE – SAMPLE')  and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Finishing') and    r.Department in ('" & _Department & "')  and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.RESCHEDULE – WASH','24.RESCHEDULE – STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RE SCHEDULE – SAMPLE')  "

                    End If
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Finishing') and    r.Merchant in ('" & _Merchant & "') and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.RESCHEDULE – WASH','24.RESCHEDULE – STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RE SCHEDULE – SAMPLE')  and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Finishing') and   r.Merchant in ('" & _Merchant & "') and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.RESCHEDULE – WASH','24.RESCHEDULE – STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RE SCHEDULE – SAMPLE')  and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Finishing') and  r.Merchant in ('" & _Merchant & "') and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.RESCHEDULE – WASH','24.RESCHEDULE – STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RE SCHEDULE – SAMPLE')  "

                    End If
                Else
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Finishing') and    r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.RESCHEDULE – WASH','24.RESCHEDULE – STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RE SCHEDULE – SAMPLE')  and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Finishing') and    r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.RESCHEDULE – WASH','24.RESCHEDULE – STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RE SCHEDULE – SAMPLE')  and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Finishing') and   r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.RESCHEDULE – WASH','24.RESCHEDULE – STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RE SCHEDULE – SAMPLE')  "

                    End If
                End If

                I = 0
                'T01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetMEETING", New SqlParameter("@cQryType", "DTY"), New SqlParameter("@vcWhereClause", vcWharer))
                'T01 = DBEngin.ExecuteDa
                Dim _DyeMC As String

                For Each DTRow3 As DataRow In T01.Tables(0).Rows
                    Dim newRow As DataRow = c_dataCustomer1.NewRow
                    Dim _Tollaranz As Double

                    'For Each DTRow1 As DataRow In M01.Tables(0).Rows
                    _DyeMC = ""

                    Dim _Material As String
                    _Material = T01.Tables(0).Rows(I)("Metrrial")
                    '_Material = Microsoft.VisualBasic.Right(_Material, 7)
                    '_Material = Microsoft.VisualBasic.Left(_Material, 2) & "-" & Microsoft.VisualBasic.Right(_Material, 5)
                    newRow("Sales Order") = T01.Tables(0).Rows(I)("Sales_Order")
                    newRow("Line Item") = T01.Tables(0).Rows(I)("Line_Item")
                    newRow("Material") = _Material
                    newRow("Description") = T01.Tables(0).Rows(I)("Met_Des")
                    newRow("Batch Qty (Mtr)") = Microsoft.VisualBasic.Format(T01.Tables(0).Rows(I)("PRD_Qty"), "#.00")
                    _QtyMtr = _QtyMtr + T01.Tables(0).Rows(I)("PRD_Qty")
                    newRow("Delivary Date") = Month(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Microsoft.VisualBasic.Day(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Year(T01.Tables(0).Rows(I)("Del_Date"))
                    _delivaryDT = Month(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Microsoft.VisualBasic.Day(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Year(T01.Tables(0).Rows(I)("Del_Date"))
                    newRow("Batch No") = T01.Tables(0).Rows(I)("Prduct_Order")
                    newRow("Customer") = T01.Tables(0).Rows(I)("Customer")

                    Sql = "select * from ZPP_DEL where Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and Line_Item='" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & "' and Delivery_Date='" & T01.Tables(0).Rows(I)("Del_Date") & "' and Product_Order='" & Trim(T01.Tables(0).Rows(I)("Prduct_Order")) & "'"
                    dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(dsUser) Then
                        newRow("Batch Qty (Kg)") = Microsoft.VisualBasic.Format(dsUser.Tables(0).Rows(0)("Order_Qty_Kg"), "#.00")
                        _QtyKg = _QtyKg + dsUser.Tables(0).Rows(0)("Order_Qty_Kg")
                        newRow("No.Of.Dys In S/L") = dsUser.Tables(0).Rows(0)("No_Day_Same_Opp")
                        newRow("NC Comment") = dsUser.Tables(0).Rows(0)("NC_Comment")
                        newRow("Next Operation") = dsUser.Tables(0).Rows(0)("Status")
                    End If

                    Sql = "select * from T07OTD_Comment1 where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07Material='" & _Material & "' and T07BatchNo='" & T01.Tables(0).Rows(I)("Prduct_Order") & "' order by T07Date DESC"
                    _OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(_OTD) Then
                        If IsDate(_OTD.Tables(0).Rows(0)("T07Posible_Date")) Then
                            If Year(_OTD.Tables(0).Rows(0)("T07Posible_Date")) = "1900" Then
                            Else
                                newRow("Possible Devivary Date") = _OTD.Tables(0).Rows(0)("T07Posible_Date")

                            End If
                        End If

                        newRow("Reason") = _OTD.Tables(0).Rows(0)("T07Reason")
                        ' newRow(_Todayupdate) = _OTD.Tables(0).Rows(0)("T07Comment")
                        newRow("LIB dep") = _OTD.Tables(0).Rows(0)("T07LIB")

                        _YesDate = Today.AddDays(-1)
                        Sql = "select * from T07OTD_Comment1 where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07BatchNo='" & T01.Tables(0).Rows(I)("Prduct_Order") & "' and T07Date<'" & Today & "' order by T07Date DESC"
                        _OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                        If isValidDataset(_OTD) Then
                            newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")
                        End If

                        Sql = "select * from FR_Update where Batch_No='" & Trim(T01.Tables(0).Rows(I)("Prduct_Order")) & "'"
                        _OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                        If isValidDataset(_OTD) Then
                            If Not DBNull.Value.Equals(_OTD.Tables(0).Rows(0)("Dye_Machine")) Then
                                _DyeMC = _OTD.Tables(0).Rows(0)("Dye_Machine")
                                newRow(_Todayupdate) = "Dye Plan on " & Month(_OTD.Tables(0).Rows(0)("Dye_Pln_Date")) & "/" & Microsoft.VisualBasic.Day(_OTD.Tables(0).Rows(0)("Dye_Pln_Date")) & "/" & Year(_OTD.Tables(0).Rows(0)("Dye_Pln_Date")) & "-" & _OTD.Tables(0).Rows(0)("Dye_Machine")
                            Else
                                newRow(_Todayupdate) = "Dye Plan on " & Month(_OTD.Tables(0).Rows(0)("Dye_Pln_Date")) & "/" & Microsoft.VisualBasic.Day(_OTD.Tables(0).Rows(0)("Dye_Pln_Date")) & "/" & Year(_OTD.Tables(0).Rows(0)("Dye_Pln_Date"))
                            End If
                        End If
                    Else
                        Sql = "select * from T07OTD_Comment1 where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07Material='" & _Material & "'  order by T07Date DESC"
                        _OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                        If isValidDataset(_OTD) Then
                            'newRow("Possible Devivary Date") = _OTD.Tables(0).Rows(0)("T07Posible_Date")
                            'newRow("Reason") = _OTD.Tables(0).Rows(0)("T07Reason")
                            If _OTD.Tables(0).Rows(0)("T07date") = Today Then

                            Else
                                '   newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")

                            End If

                            _YesDate = Today.AddDays(-1)
                            Sql = "select * from T07OTD_Comment1 where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07BatchNo='" & T01.Tables(0).Rows(I)("Prduct_Order") & "' and T07Date<'" & Today & "' order by T07Date DESC"
                            _OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                            If isValidDataset(_OTD) Then
                                newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")
                            End If

                            'newRow("LIB dep") = _OTD.Tables(0).Rows(0)("T07LIB")
                        Else
                            Sql = "select * from FR_Update where Batch_No='" & Trim(T01.Tables(0).Rows(I)("Prduct_Order")) & "'"
                            _OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                            If isValidDataset(_OTD) Then
                                If Not DBNull.Value.Equals(_OTD.Tables(0).Rows(0)("Dye_Machine")) Then
                                    _DyeMC = _OTD.Tables(0).Rows(0)("Dye_Machine")
                                    newRow(_Todayupdate) = "Dye Plan on " & Month(_OTD.Tables(0).Rows(0)("Dye_Pln_Date")) & "/" & Microsoft.VisualBasic.Day(_OTD.Tables(0).Rows(0)("Dye_Pln_Date")) & "/" & Year(_OTD.Tables(0).Rows(0)("Dye_Pln_Date")) & "-" & _OTD.Tables(0).Rows(0)("Dye_Machine")
                                Else
                                    newRow(_Todayupdate) = "Dye Plan on " & Month(_OTD.Tables(0).Rows(0)("Dye_Pln_Date")) & "/" & Microsoft.VisualBasic.Day(_OTD.Tables(0).Rows(0)("Dye_Pln_Date")) & "/" & Year(_OTD.Tables(0).Rows(0)("Dye_Pln_Date"))
                                End If
                            End If
                        End If
                    End If

                    newRow("Week") = "Week " & DatePart(DateInterval.WeekOfYear, _delivaryDT)
                    c_dataCustomer1.Rows.Add(newRow)
                    ' End If
                    Dim _Rcount As Integer
                    _Rcount = UltraGrid4.Rows.Count
                    If Microsoft.VisualBasic.Left(_DyeMC, 1) = "S" Or Microsoft.VisualBasic.Left(_DyeMC, 1) = "Y" Or Microsoft.VisualBasic.Left(_DyeMC, 1) = "W" Or IsNumeric(_DyeMC) Then
                        UltraGrid4.Rows(_Rcount - 1).Cells(0).Appearance.BackColor = Color.IndianRed
                        UltraGrid4.Rows(_Rcount - 1).Cells(1).Appearance.BackColor = Color.IndianRed
                        UltraGrid4.Rows(_Rcount - 1).Cells(2).Appearance.BackColor = Color.IndianRed
                        UltraGrid4.Rows(_Rcount - 1).Cells(3).Appearance.BackColor = Color.IndianRed
                        UltraGrid4.Rows(_Rcount - 1).Cells(4).Appearance.BackColor = Color.IndianRed
                        UltraGrid4.Rows(_Rcount - 1).Cells(5).Appearance.BackColor = Color.IndianRed
                        UltraGrid4.Rows(_Rcount - 1).Cells(6).Appearance.BackColor = Color.IndianRed
                        UltraGrid4.Rows(_Rcount - 1).Cells(7).Appearance.BackColor = Color.IndianRed
                        UltraGrid4.Rows(_Rcount - 1).Cells(8).Appearance.BackColor = Color.IndianRed
                        UltraGrid4.Rows(_Rcount - 1).Cells(9).Appearance.BackColor = Color.IndianRed
                        UltraGrid4.Rows(_Rcount - 1).Cells(10).Appearance.BackColor = Color.IndianRed
                        UltraGrid4.Rows(_Rcount - 1).Cells(11).Appearance.BackColor = Color.IndianRed
                        UltraGrid4.Rows(_Rcount - 1).Cells(12).Appearance.BackColor = Color.IndianRed
                        UltraGrid4.Rows(_Rcount - 1).Cells(13).Appearance.BackColor = Color.IndianRed
                        UltraGrid4.Rows(_Rcount - 1).Cells(14).Appearance.BackColor = Color.IndianRed
                        UltraGrid4.Rows(_Rcount - 1).Cells(15).Appearance.BackColor = Color.IndianRed
                        UltraGrid4.Rows(_Rcount - 1).Cells(16).Appearance.BackColor = Color.IndianRed
                    End If
                    ' If Not x = M01.Tables(0).Rows.Count - 1 Then
                    ' rsT24RecHeader.MoveNext()
                    _Rowcount = _Rowcount + 1
                    I = I + 1
                Next

            ElseIf Trim(cboPO.Text) = "Aw for 2062 location" Then
                Dim _OTD As DataSet

                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2062') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')   and left(r.Met_Des,8)='" & cboQuality.Text & "' and r.NC_Comment in ('13.HELD IN OTHER REASON KNITTING','')"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2062') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')   and left(r.Met_Des,5)='" & cboQuality.Text & "' and r.NC_Comment in ('13.HELD IN OTHER REASON KNITTING','')"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2062') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.NC_Comment in ('13.HELD IN OTHER REASON KNITTING','')"

                    End If
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2062') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')    and left(r.Met_Des,8)='" & cboQuality.Text & "' and r.NC_Comment in ('13.HELD IN OTHER REASON KNITTING','')"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2062') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')   and left(r.Met_Des,5)='" & cboQuality.Text & "' and r.NC_Comment in ('13.HELD IN OTHER REASON KNITTING','')"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2062') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')and r.NC_Comment in ('13.HELD IN OTHER REASON KNITTING','') "

                    End If
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2062') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "')   and left(r.Met_Des,8)='" & cboQuality.Text & "' and r.NC_Comment in ('13.HELD IN OTHER REASON KNITTING','')"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2062') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "')   and left(r.Met_Des,5)='" & cboQuality.Text & "' and r.NC_Comment in ('13.HELD IN OTHER REASON KNITTING','')"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2062') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "') and r.NC_Comment in ('13.HELD IN OTHER REASON KNITTING','')"

                    End If
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then

                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2062') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')   and left(r.Met_Des,8)='" & cboQuality.Text & "' and r.NC_Comment in ('13.HELD IN OTHER REASON KNITTING','')"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2062') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')   and left(r.Met_Des,5)='" & cboQuality.Text & "' and r.NC_Comment in ('13.HELD IN OTHER REASON KNITTING','')"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2062') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.NC_Comment in ('13.HELD IN OTHER REASON KNITTING','')"

                    End If

                ElseIf Trim(txtCustomer.Text) <> "" Then

                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2062') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "')   and left(r.Met_Des,8)='" & cboQuality.Text & "' and r.NC_Comment in ('13.HELD IN OTHER REASON KNITTING','')"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2062') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "')   and left(r.Met_Des,5)='" & cboQuality.Text & "' and r.NC_Comment in ('13.HELD IN OTHER REASON KNITTING','')"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2062') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "')"

                    End If

                ElseIf Trim(txtDepartment.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2062') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Department in ('" & _Department & "')    and left(r.Met_Des,8)='" & cboQuality.Text & "' and r.NC_Comment in ('13.HELD IN OTHER REASON KNITTING','')"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2062') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Department in ('" & _Department & "')   and left(r.Met_Des,5)='" & cboQuality.Text & "' and r.NC_Comment in ('13.HELD IN OTHER REASON KNITTING','')"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2062') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Department in ('" & _Department & "') and r.NC_Comment in ('13.HELD IN OTHER REASON KNITTING','')"

                    End If
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2062') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Merchant in ('" & _Merchant & "')   and left(r.Met_Des,8)='" & cboQuality.Text & "' and r.NC_Comment in ('13.HELD IN OTHER REASON KNITTING','')"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2062') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Merchant in ('" & _Merchant & "')   and left(r.Met_Des,5)='" & cboQuality.Text & "' and r.NC_Comment in ('13.HELD IN OTHER REASON KNITTING','')"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2062') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Merchant in ('" & _Merchant & "') and r.NC_Comment in ('13.HELD IN OTHER REASON KNITTING','')"

                    End If
                Else
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2062') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and left(r.Met_Des,8)='" & cboQuality.Text & "' and r.NC_Comment in ('13.HELD IN OTHER REASON KNITTING','')"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2062') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and left(r.Met_Des,5)='" & cboQuality.Text & "' and r.NC_Comment in ('13.HELD IN OTHER REASON KNITTING','')"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2062') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.NC_Comment in ('13.HELD IN OTHER REASON KNITTING','')"

                    End If
                End If

                I = 0
                'T01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetMEETING", New SqlParameter("@cQryType", "DTY"), New SqlParameter("@vcWhereClause", vcWharer))
                'T01 = DBEngin.ExecuteDa
                _QtyKg = 0
                _QtyMtr = 0
                _Rowcount = 0
                For Each DTRow3 As DataRow In T01.Tables(0).Rows
                    Dim newRow As DataRow = c_dataCustomer1.NewRow
                    Dim _Tollaranz As Double

                    'For Each DTRow1 As DataRow In M01.Tables(0).Rows


                    Dim _Material As String
                    _Material = T01.Tables(0).Rows(I)("Metrrial")
                    '_Material = Microsoft.VisualBasic.Right(_Material, 7)
                    '_Material = Microsoft.VisualBasic.Left(_Material, 2) & "-" & Microsoft.VisualBasic.Right(_Material, 5)
                    newRow("Sales Order") = T01.Tables(0).Rows(I)("Sales_Order")
                    newRow("Line Item") = T01.Tables(0).Rows(I)("Line_Item")
                    newRow("Material") = _Material
                    newRow("Description") = T01.Tables(0).Rows(I)("Met_Des")
                    newRow("Batch Qty (Mtr)") = Microsoft.VisualBasic.Format(T01.Tables(0).Rows(I)("PRD_Qty"), "#.00")
                    _QtyMtr = _QtyMtr + T01.Tables(0).Rows(I)("PRD_Qty")
                    newRow("Delivary Date") = Month(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Microsoft.VisualBasic.Day(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Year(T01.Tables(0).Rows(I)("Del_Date"))
                    _delivaryDT = Month(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Microsoft.VisualBasic.Day(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Year(T01.Tables(0).Rows(I)("Del_Date"))
                    newRow("Batch No") = T01.Tables(0).Rows(I)("Prduct_Order")
                    newRow("Customer") = T01.Tables(0).Rows(I)("Customer")

                    Sql = "select * from ZPP_DEL where Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and Line_Item='" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & "' and Delivery_Date='" & T01.Tables(0).Rows(I)("Del_Date") & "' and Product_Order='" & Trim(T01.Tables(0).Rows(I)("Prduct_Order")) & "'"
                    dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(dsUser) Then
                        newRow("Batch Qty (Kg)") = Microsoft.VisualBasic.Format(dsUser.Tables(0).Rows(0)("Order_Qty_Kg"), "#.00")
                        _QtyKg = _QtyKg + dsUser.Tables(0).Rows(0)("Order_Qty_Kg")
                        newRow("No.Of.Dys In S/L") = dsUser.Tables(0).Rows(0)("No_Day_Same_Opp")
                        newRow("NC Comment") = dsUser.Tables(0).Rows(0)("NC_Comment")
                        '  newRow("Next Operation") = dsUser.Tables(0).Rows(0)("Status")
                    End If

                    Sql = "select * from T07OTD_Comment1 where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07Material='" & _Material & "' and T07BatchNo='" & T01.Tables(0).Rows(I)("Prduct_Order") & "' order by T07Date DESC"
                    _OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(_OTD) Then
                        If IsDate(_OTD.Tables(0).Rows(0)("T07Posible_Date")) Then
                            If Year(_OTD.Tables(0).Rows(0)("T07Posible_Date")) = "1900" Then
                            Else
                                newRow("Possible Devivary Date") = _OTD.Tables(0).Rows(0)("T07Posible_Date")

                            End If
                        End If

                        newRow("Reason") = _OTD.Tables(0).Rows(0)("T07Reason")
                        If _OTD.Tables(0).Rows(0)("T07date") = Today Then
                            newRow(_Todayupdate) = _OTD.Tables(0).Rows(0)("T07Comment")
                        Else
                            newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")
                        End If
                        newRow("LIB dep") = _OTD.Tables(0).Rows(0)("T07LIB")

                        _YesDate = Today.AddDays(-1)
                        Sql = "select * from T07OTD_Comment1 where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07BatchNo='" & T01.Tables(0).Rows(I)("Prduct_Order") & "' and T07Date<'" & Today & "' order by T07Date DESC"
                        _OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                        If isValidDataset(_OTD) Then
                            newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")
                        End If

                    Else
                        'Sql = "select * from T07OTD_Comment where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07Material='" & _Material & "'  order by T07Date DESC"
                        '_OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                        'If isValidDataset(_OTD) Then
                        '    'newRow("Possible Devivary Date") = _OTD.Tables(0).Rows(0)("T07Posible_Date")
                        '    'newRow("Reason") = _OTD.Tables(0).Rows(0)("T07Reason")
                        '    newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")
                        '    'newRow("LIB dep") = _OTD.Tables(0).Rows(0)("T07LIB")
                        'End If
                    End If
                    newRow("Week") = "Week " & DatePart(DateInterval.WeekOfYear, _delivaryDT)
                    c_dataCustomer1.Rows.Add(newRow)
                    ' End If

                    ' If Not x = M01.Tables(0).Rows.Count - 1 Then
                    ' rsT24RecHeader.MoveNext()
                    _Rowcount = _Rowcount + 1
                    I = I + 1
                Next

                'NEW LOCATION DEVELOPING DATE 2016.6.14
                'REQUEST BY AMILA
            ElseIf Trim(cboPO.Text) = "Aw for 2059 location" Then
                Dim _OTD As DataSet

                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2059') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')   and left(r.Met_Des,8)='" & cboQuality.Text & "' and r.NC_Comment in ('13.HELD IN OTHER REASON KNITTING','')"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2059') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')   and left(r.Met_Des,5)='" & cboQuality.Text & "' and r.NC_Comment in ('13.HELD IN OTHER REASON KNITTING','')"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2059') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.NC_Comment in ('13.HELD IN OTHER REASON KNITTING','')"

                    End If
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2059') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')    and left(r.Met_Des,8)='" & cboQuality.Text & "' and r.NC_Comment in ('13.HELD IN OTHER REASON KNITTING','')"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2059') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')   and left(r.Met_Des,5)='" & cboQuality.Text & "' and r.NC_Comment in ('13.HELD IN OTHER REASON KNITTING','')"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2059') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')and r.NC_Comment in ('13.HELD IN OTHER REASON KNITTING','') "

                    End If
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2059') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "')   and left(r.Met_Des,8)='" & cboQuality.Text & "' and r.NC_Comment in ('13.HELD IN OTHER REASON KNITTING','')"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2059') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "')   and left(r.Met_Des,5)='" & cboQuality.Text & "' and r.NC_Comment in ('13.HELD IN OTHER REASON KNITTING','')"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2059') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "') and r.NC_Comment in ('13.HELD IN OTHER REASON KNITTING','')"

                    End If
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then

                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2059') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')   and left(r.Met_Des,8)='" & cboQuality.Text & "' and r.NC_Comment in ('13.HELD IN OTHER REASON KNITTING','')"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2059') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')   and left(r.Met_Des,5)='" & cboQuality.Text & "' and r.NC_Comment in ('13.HELD IN OTHER REASON KNITTING','')"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2059') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.NC_Comment in ('13.HELD IN OTHER REASON KNITTING','')"

                    End If

                ElseIf Trim(txtCustomer.Text) <> "" Then

                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2059') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "')   and left(r.Met_Des,8)='" & cboQuality.Text & "' and r.NC_Comment in ('13.HELD IN OTHER REASON KNITTING','')"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2059') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "')   and left(r.Met_Des,5)='" & cboQuality.Text & "' and r.NC_Comment in ('13.HELD IN OTHER REASON KNITTING','')"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2059') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "')"

                    End If

                ElseIf Trim(txtDepartment.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2059') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Department in ('" & _Department & "')    and left(r.Met_Des,8)='" & cboQuality.Text & "' and r.NC_Comment in ('13.HELD IN OTHER REASON KNITTING','')"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2059') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Department in ('" & _Department & "')   and left(r.Met_Des,5)='" & cboQuality.Text & "' and r.NC_Comment in ('13.HELD IN OTHER REASON KNITTING','')"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2059') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Department in ('" & _Department & "') and r.NC_Comment in ('13.HELD IN OTHER REASON KNITTING','')"

                    End If
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2059') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Merchant in ('" & _Merchant & "')   and left(r.Met_Des,8)='" & cboQuality.Text & "' and r.NC_Comment in ('13.HELD IN OTHER REASON KNITTING','')"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2059') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Merchant in ('" & _Merchant & "')   and left(r.Met_Des,5)='" & cboQuality.Text & "' and r.NC_Comment in ('13.HELD IN OTHER REASON KNITTING','')"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2059') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Merchant in ('" & _Merchant & "') and r.NC_Comment in ('13.HELD IN OTHER REASON KNITTING','')"

                    End If
                Else
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2059') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and left(r.Met_Des,8)='" & cboQuality.Text & "' and r.NC_Comment in ('13.HELD IN OTHER REASON KNITTING','')"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2059') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and left(r.Met_Des,5)='" & cboQuality.Text & "' and r.NC_Comment in ('13.HELD IN OTHER REASON KNITTING','')"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2059') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.NC_Comment in ('13.HELD IN OTHER REASON KNITTING','')"

                    End If
                End If

                I = 0
                'T01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetMEETING", New SqlParameter("@cQryType", "DTY"), New SqlParameter("@vcWhereClause", vcWharer))
                'T01 = DBEngin.ExecuteDa
                _QtyKg = 0
                _QtyMtr = 0
                _Rowcount = 0
                For Each DTRow3 As DataRow In T01.Tables(0).Rows
                    Dim newRow As DataRow = c_dataCustomer1.NewRow
                    Dim _Tollaranz As Double

                    'For Each DTRow1 As DataRow In M01.Tables(0).Rows


                    Dim _Material As String
                    _Material = T01.Tables(0).Rows(I)("Metrrial")
                    '_Material = Microsoft.VisualBasic.Right(_Material, 7)
                    '_Material = Microsoft.VisualBasic.Left(_Material, 2) & "-" & Microsoft.VisualBasic.Right(_Material, 5)
                    newRow("Sales Order") = T01.Tables(0).Rows(I)("Sales_Order")
                    newRow("Line Item") = T01.Tables(0).Rows(I)("Line_Item")
                    newRow("Material") = _Material
                    newRow("Description") = T01.Tables(0).Rows(I)("Met_Des")
                    newRow("Batch Qty (Mtr)") = Microsoft.VisualBasic.Format(T01.Tables(0).Rows(I)("PRD_Qty"), "#.00")
                    _QtyMtr = _QtyMtr + T01.Tables(0).Rows(I)("PRD_Qty")
                    newRow("Delivary Date") = Month(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Microsoft.VisualBasic.Day(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Year(T01.Tables(0).Rows(I)("Del_Date"))
                    _delivaryDT = Month(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Microsoft.VisualBasic.Day(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Year(T01.Tables(0).Rows(I)("Del_Date"))
                    newRow("Batch No") = T01.Tables(0).Rows(I)("Prduct_Order")
                    newRow("Customer") = T01.Tables(0).Rows(I)("Customer")

                    Sql = "select * from ZPP_DEL where Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and Line_Item='" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & "' and Delivery_Date='" & T01.Tables(0).Rows(I)("Del_Date") & "' and Product_Order='" & Trim(T01.Tables(0).Rows(I)("Prduct_Order")) & "'"
                    dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(dsUser) Then
                        newRow("Batch Qty (Kg)") = Microsoft.VisualBasic.Format(dsUser.Tables(0).Rows(0)("Order_Qty_Kg"), "#.00")
                        _QtyKg = _QtyKg + dsUser.Tables(0).Rows(0)("Order_Qty_Kg")
                        newRow("No.Of.Dys In S/L") = dsUser.Tables(0).Rows(0)("No_Day_Same_Opp")
                        newRow("NC Comment") = dsUser.Tables(0).Rows(0)("NC_Comment")
                        '  newRow("Next Operation") = dsUser.Tables(0).Rows(0)("Status")
                    End If

                    Sql = "select * from T07OTD_Comment1 where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07Material='" & _Material & "' and T07BatchNo='" & T01.Tables(0).Rows(I)("Prduct_Order") & "' order by T07Date DESC"
                    _OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(_OTD) Then
                        If IsDate(_OTD.Tables(0).Rows(0)("T07Posible_Date")) Then
                            If Year(_OTD.Tables(0).Rows(0)("T07Posible_Date")) = "1900" Then
                            Else
                                newRow("Possible Devivary Date") = _OTD.Tables(0).Rows(0)("T07Posible_Date")

                            End If
                        End If

                        newRow("Reason") = _OTD.Tables(0).Rows(0)("T07Reason")
                        If _OTD.Tables(0).Rows(0)("T07date") = Today Then
                            newRow(_Todayupdate) = _OTD.Tables(0).Rows(0)("T07Comment")
                        Else
                            newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")
                        End If
                        newRow("LIB dep") = _OTD.Tables(0).Rows(0)("T07LIB")

                        _YesDate = Today.AddDays(-1)
                        Sql = "select * from T07OTD_Comment1 where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07BatchNo='" & T01.Tables(0).Rows(I)("Prduct_Order") & "' and T07Date<'" & Today & "' order by T07Date DESC"
                        _OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                        If isValidDataset(_OTD) Then
                            newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")
                        End If

                    Else
                        'Sql = "select * from T07OTD_Comment where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07Material='" & _Material & "'  order by T07Date DESC"
                        '_OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                        'If isValidDataset(_OTD) Then
                        '    'newRow("Possible Devivary Date") = _OTD.Tables(0).Rows(0)("T07Posible_Date")
                        '    'newRow("Reason") = _OTD.Tables(0).Rows(0)("T07Reason")
                        '    newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")
                        '    'newRow("LIB dep") = _OTD.Tables(0).Rows(0)("T07LIB")
                        'End If
                    End If
                    newRow("Week") = "Week " & DatePart(DateInterval.WeekOfYear, _delivaryDT)
                    c_dataCustomer1.Rows.Add(newRow)
                    ' End If

                    ' If Not x = M01.Tables(0).Rows.Count - 1 Then
                    ' rsT24RecHeader.MoveNext()
                    _Rowcount = _Rowcount + 1
                    I = I + 1
                Next

            ElseIf Trim(cboPO.Text) = "Aw for 2065 location" Then
                Dim _OTD As DataSet
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')   and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')   and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')"

                    End If
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')    and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')   and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') "

                    End If
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "')   and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "')   and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "')"

                    End If
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then

                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')   and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')   and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')"

                    End If

                ElseIf Trim(txtCustomer.Text) <> "" Then

                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "')   and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "')   and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "')"

                    End If

                ElseIf Trim(txtDepartment.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Department in ('" & _Department & "')    and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Department in ('" & _Department & "')   and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Department in ('" & _Department & "') "

                    End If
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Merchant in ('" & _Merchant & "')   and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Merchant in ('" & _Merchant & "')   and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Merchant in ('" & _Merchant & "')"

                    End If
                Else
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') "

                    End If
                End If

                I = 0
                'T01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetMEETING", New SqlParameter("@cQryType", "DTY"), New SqlParameter("@vcWhereClause", vcWharer))
                'T01 = DBEngin.ExecuteDa
                _Rowcount = 0
                _QtyKg = 0
                _QtyMtr = 0
                For Each DTRow3 As DataRow In T01.Tables(0).Rows
                    Dim newRow As DataRow = c_dataCustomer1.NewRow
                    Dim _Tollaranz As Double

                    'For Each DTRow1 As DataRow In M01.Tables(0).Rows


                    Dim _Material As String
                    _Material = T01.Tables(0).Rows(I)("Metrrial")
                    '_Material = Microsoft.VisualBasic.Right(_Material, 7)
                    '_Material = Microsoft.VisualBasic.Left(_Material, 2) & "-" & Microsoft.VisualBasic.Right(_Material, 5)
                    newRow("Sales Order") = T01.Tables(0).Rows(I)("Sales_Order")
                    newRow("Line Item") = T01.Tables(0).Rows(I)("Line_Item")
                    newRow("Material") = _Material
                    newRow("Description") = T01.Tables(0).Rows(I)("Met_Des")
                    newRow("Batch Qty (Mtr)") = Microsoft.VisualBasic.Format(T01.Tables(0).Rows(I)("PRD_Qty"), "#.00")
                    _QtyMtr = _QtyMtr + T01.Tables(0).Rows(I)("PRD_Qty")
                    newRow("Delivary Date") = Month(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Microsoft.VisualBasic.Day(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Year(T01.Tables(0).Rows(I)("Del_Date"))
                    _delivaryDT = Month(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Microsoft.VisualBasic.Day(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Year(T01.Tables(0).Rows(I)("Del_Date"))
                    newRow("Batch No") = T01.Tables(0).Rows(I)("Prduct_Order")
                    newRow("Customer") = T01.Tables(0).Rows(I)("Customer")

                    Sql = "select * from ZPP_DEL where Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and Line_Item='" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & "' and Delivery_Date='" & T01.Tables(0).Rows(I)("Del_Date") & "' and Product_Order='" & Trim(T01.Tables(0).Rows(I)("Prduct_Order")) & "'"
                    dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(dsUser) Then
                        newRow("Batch Qty (Kg)") = Microsoft.VisualBasic.Format(dsUser.Tables(0).Rows(0)("Order_Qty_Kg"), "#.00")
                        _QtyKg = _QtyKg + dsUser.Tables(0).Rows(0)("Order_Qty_Kg")
                        newRow("No.Of.Dys In S/L") = dsUser.Tables(0).Rows(0)("No_Day_Same_Opp")
                        newRow("NC Comment") = dsUser.Tables(0).Rows(0)("NC_Comment")
                        ' newRow("Next Operation") = dsUser.Tables(0).Rows(0)("Status")
                    End If

                    Sql = "select * from T07OTD_Comment1 where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07Material='" & _Material & "' and T07BatchNo='" & T01.Tables(0).Rows(I)("Prduct_Order") & "' order by T07Date DESC"
                    _OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(_OTD) Then
                        If IsDate(_OTD.Tables(0).Rows(0)("T07Posible_Date")) Then
                            If Year(_OTD.Tables(0).Rows(0)("T07Posible_Date")) = "1900" Then
                            Else
                                newRow("Possible Devivary Date") = _OTD.Tables(0).Rows(0)("T07Posible_Date")

                            End If
                        End If

                        newRow("Reason") = _OTD.Tables(0).Rows(0)("T07Reason")

                        If _OTD.Tables(0).Rows(0)("T07date") = Today Then
                            newRow(_Todayupdate) = _OTD.Tables(0).Rows(0)("T07Comment")
                        Else
                            newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")
                        End If
                        newRow("LIB dep") = _OTD.Tables(0).Rows(0)("T07LIB")

                        _YesDate = Today.AddDays(-1)
                        Sql = "select * from T07OTD_Comment1 where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07BatchNo='" & T01.Tables(0).Rows(I)("Prduct_Order") & "' and T07Date<'" & Today & "' order by T07Date DESC"
                        _OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                        If isValidDataset(_OTD) Then
                            newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")
                        End If

                    Else
                        'Sql = "select * from T07OTD_Comment where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07Material='" & _Material & "'  order by T07Date DESC"
                        '_OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                        'If isValidDataset(_OTD) Then
                        '    'newRow("Possible Devivary Date") = _OTD.Tables(0).Rows(0)("T07Posible_Date")
                        '    'newRow("Reason") = _OTD.Tables(0).Rows(0)("T07Reason")
                        '    newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")
                        '    'newRow("LIB dep") = _OTD.Tables(0).Rows(0)("T07LIB")
                        'End If
                    End If
                    newRow("Week") = "Week " & DatePart(DateInterval.WeekOfYear, _delivaryDT)
                    c_dataCustomer1.Rows.Add(newRow)
                    ' End If

                    ' If Not x = M01.Tables(0).Rows.Count - 1 Then
                    ' rsT24RecHeader.MoveNext()
                    _Rowcount = _Rowcount + 1
                    I = I + 1
                Next
            ElseIf Trim(cboPO.Text) = "Aw Pigment" Then
                Dim _OTD As DataSet
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'   and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  and left(r.Met_Des,8)='" & cboQuality.Text & "' and r.NC_Comment in ('9.Aw pad Pigment','10.AW PAD UV','9.AW PAD PIGMENT')"
                        Else
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'   and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  and left(r.Met_Des,5)='" & cboQuality.Text & "' and r.NC_Comment in ('9.Aw pad Pigment','10.AW PAD UV','9.AW PAD PIGMENT')"
                        End If
                    Else
                        vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'   and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')and r.NC_Comment in ('9.Aw pad Pigment','10.AW PAD UV','9.AW PAD PIGMENT') "

                    End If
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'   and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')   and left(r.Met_Des,8)='" & cboQuality.Text & "' and r.NC_Comment in ('9.Aw pad Pigment','10.AW PAD UV','9.AW PAD PIGMENT')"
                        Else
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')   and left(r.Met_Des,5)='" & cboQuality.Text & "' and r.NC_Comment in ('9.Aw pad Pigment','10.AW PAD UV','9.AW PAD PIGMENT')"
                        End If
                    Else
                        vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.NC_Comment in ('9.Aw pad Pigment','10.AW PAD UV','9.AW PAD PIGMENT') "

                    End If
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'   and  Location='Finishing' and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "')  and left(r.Met_Des,8)='" & cboQuality.Text & "' and r.NC_Comment in ('9.Aw pad Pigment','10.AW PAD UV','9.AW PAD PIGMENT')"
                        Else
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'   and  Location='Finishing' and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "')  and left(r.Met_Des,5)='" & cboQuality.Text & "' and r.NC_Comment in ('9.Aw pad Pigment','10.AW PAD UV','9.AW PAD PIGMENT')"
                        End If
                    Else
                        vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'   and  Location='Finishing' and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "') and r.NC_Comment in ('9.Aw pad Pigment','10.AW PAD UV','9.AW PAD PIGMENT')"

                    End If
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then

                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'   and  Location='Finishing'  and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  and left(r.Met_Des,8)='" & cboQuality.Text & "' and r.NC_Comment in ('9.Aw pad Pigment','10.AW PAD UV','9.AW PAD PIGMENT')"
                        Else
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'   and  Location='Finishing' and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  and left(r.Met_Des,5)='" & cboQuality.Text & "' and r.NC_Comment in ('9.Aw pad Pigment','10.AW PAD UV','9.AW PAD PIGMENT')"
                        End If
                    Else
                        vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'   and  Location='Finishing'  and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.NC_Comment in ('9.Aw pad Pigment','10.AW PAD UV','9.AW PAD PIGMENT')"

                    End If

                ElseIf Trim(txtCustomer.Text) <> "" Then

                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'  and  Location='Finishing' and s.Customer in ('" & _Customer & "')   and left(r.Met_Des,8)='" & cboQuality.Text & "' and r.NC_Comment in ('9.Aw pad Pigment','10.AW PAD UV','9.AW PAD PIGMENT')"
                        Else
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and  Location='Finishing' and s.Customer in ('" & _Customer & "')  and left(r.Met_Des,5)='" & cboQuality.Text & "' and r.NC_Comment in ('9.Aw pad Pigment','10.AW PAD UV','9.AW PAD PIGMENT')"
                        End If
                    Else
                        vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.NC_Comment in ('9.Aw pad Pigment','10.AW PAD UV','9.AW PAD PIGMENT') "

                    End If

                ElseIf Trim(txtDepartment.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'  and  Location='Finishing'  and r.Department in ('" & _Department & "')   and left(r.Met_Des,8)='" & cboQuality.Text & "' and r.NC_Comment in ('9.Aw pad Pigment','10.AW PAD UV','9.AW PAD PIGMENT')"
                        Else
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'  and  Location='Finishing'  and r.Department in ('" & _Department & "')  and left(r.Met_Des,5)='" & cboQuality.Text & "' and r.NC_Comment in ('9.Aw pad Pigment','10.AW PAD UV','9.AW PAD PIGMENT')"
                        End If
                    Else
                        vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'  and  Location='Finishing' and  and r.Department in ('" & _Department & "') and r.NC_Comment in ('9.Aw pad Pigment','10.AW PAD UV','9.AW PAD PIGMENT') "

                    End If
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'   and  Location='Finishing'  and r.Merchant in ('" & _Merchant & "')  and left(r.Met_Des,8)='" & cboQuality.Text & "' and r.NC_Comment in ('9.Aw pad Pigment','10.AW PAD UV','9.AW PAD PIGMENT')"
                        Else
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'   and  Location='Finishing'  and r.Merchant in ('" & _Merchant & "')  and left(r.Met_Des,5)='" & cboQuality.Text & "' and r.NC_Comment in ('9.Aw pad Pigment','10.AW PAD UV','9.AW PAD PIGMENT')"
                        End If
                    Else
                        vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'  and  Location='Finishing' and  r.Merchant in ('" & _Merchant & "') and r.NC_Comment in ('9.Aw pad Pigment','10.AW PAD UV','9.AW PAD PIGMENT')"

                    End If
                Else
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'  and  Location='Finishing'  and left(r.Met_Des,8)='" & cboQuality.Text & "' and r.NC_Comment in ('9.Aw pad Pigment','10.AW PAD UV','9.AW PAD PIGMENT')"
                        Else
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'   and  Location='Finishing'   and left(r.Met_Des,5)='" & cboQuality.Text & "' and r.NC_Comment in ('9.Aw pad Pigment','10.AW PAD UV','9.AW PAD PIGMENT')"
                        End If
                    Else
                        vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'   and  Location='Finishing' and r.NC_Comment in ('9.Aw pad Pigment','9.AW PAD PIGMENT','10.AW PAD UV') "

                    End If
                End If

                I = 0
                'T01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetMEETING", New SqlParameter("@cQryType", "DTY"), New SqlParameter("@vcWhereClause", vcWharer))
                'T01 = DBEngin.ExecuteDa
                _QtyKg = 0
                _QtyMtr = 0
                _Rowcount = 0
                For Each DTRow3 As DataRow In T01.Tables(0).Rows
                    Dim newRow As DataRow = c_dataCustomer1.NewRow
                    Dim _Tollaranz As Double

                    'For Each DTRow1 As DataRow In M01.Tables(0).Rows


                    Dim _Material As String
                    _Material = T01.Tables(0).Rows(I)("Metrrial")
                    '_Material = Microsoft.VisualBasic.Right(_Material, 7)
                    '_Material = Microsoft.VisualBasic.Left(_Material, 2) & "-" & Microsoft.VisualBasic.Right(_Material, 5)
                    newRow("Sales Order") = T01.Tables(0).Rows(I)("Sales_Order")
                    newRow("Line Item") = T01.Tables(0).Rows(I)("Line_Item")
                    newRow("Material") = _Material
                    newRow("Description") = T01.Tables(0).Rows(I)("Met_Des")
                    newRow("Batch Qty (Mtr)") = Microsoft.VisualBasic.Format(T01.Tables(0).Rows(I)("PRD_Qty"), "#.00")
                    _QtyMtr = _QtyMtr + T01.Tables(0).Rows(I)("PRD_Qty")
                    newRow("Delivary Date") = Month(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Microsoft.VisualBasic.Day(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Year(T01.Tables(0).Rows(I)("Del_Date"))
                    _delivaryDT = Month(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Microsoft.VisualBasic.Day(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Year(T01.Tables(0).Rows(I)("Del_Date"))
                    newRow("Batch No") = T01.Tables(0).Rows(I)("Prduct_Order")
                    newRow("Customer") = T01.Tables(0).Rows(I)("Customer")

                    Sql = "select * from ZPP_DEL where Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and Line_Item='" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & "' and Delivery_Date='" & T01.Tables(0).Rows(I)("Del_Date") & "' and Product_Order='" & Trim(T01.Tables(0).Rows(I)("Prduct_Order")) & "'"
                    dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(dsUser) Then
                        newRow("Batch Qty (Kg)") = Microsoft.VisualBasic.Format(dsUser.Tables(0).Rows(0)("Order_Qty_Kg"), "#.00")
                        _QtyKg = _QtyKg + dsUser.Tables(0).Rows(0)("Order_Qty_Kg")
                        newRow("No.Of.Dys In S/L") = dsUser.Tables(0).Rows(0)("No_Day_Same_Opp")
                        newRow("NC Comment") = dsUser.Tables(0).Rows(0)("NC_Comment")
                        newRow("Next Operation") = dsUser.Tables(0).Rows(0)("Status")
                    End If

                    Sql = "select * from T07OTD_Comment1 where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07Material='" & _Material & "' and T07BatchNo='" & T01.Tables(0).Rows(I)("Prduct_Order") & "' order by T07Date DESC"
                    _OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(_OTD) Then
                        If IsDate(_OTD.Tables(0).Rows(0)("T07Posible_Date")) Then
                            If Year(_OTD.Tables(0).Rows(0)("T07Posible_Date")) = "1900" Then
                            Else
                                newRow("Possible Devivary Date") = _OTD.Tables(0).Rows(0)("T07Posible_Date")

                            End If
                        End If

                        newRow("Reason") = _OTD.Tables(0).Rows(0)("T07Reason")
                        If _OTD.Tables(0).Rows(0)("T07date") = Today Then
                            newRow(_Todayupdate) = _OTD.Tables(0).Rows(0)("T07Comment")
                        Else
                            newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")
                        End If
                        newRow("LIB dep") = _OTD.Tables(0).Rows(0)("T07LIB")

                        _YesDate = Today.AddDays(-1)
                        Sql = "select * from T07OTD_Comment1 where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07BatchNo='" & T01.Tables(0).Rows(I)("Prduct_Order") & "' and T07Date<'" & Today & "' order by T07Date DESC"
                        _OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                        If isValidDataset(_OTD) Then
                            newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")
                        End If

                    Else
                        'Sql = "select * from T07OTD_Comment where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07Material='" & _Material & "'  order by T07Date DESC"
                        '_OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                        'If isValidDataset(_OTD) Then
                        '    'newRow("Possible Devivary Date") = _OTD.Tables(0).Rows(0)("T07Posible_Date")
                        '    'newRow("Reason") = _OTD.Tables(0).Rows(0)("T07Reason")
                        '    newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")
                        '    'newRow("LIB dep") = _OTD.Tables(0).Rows(0)("T07LIB")
                        'End If
                    End If

                    newRow("Week") = "Week " & DatePart(DateInterval.WeekOfYear, _delivaryDT)
                    c_dataCustomer1.Rows.Add(newRow)
                    ' End If

                    ' If Not x = M01.Tables(0).Rows.Count - 1 Then
                    ' rsT24RecHeader.MoveNext()
                    _Rowcount = _Rowcount + 1
                    I = I + 1
                Next
            ElseIf Trim(cboPO.Text) = "Pilot & AW Shade comment" Then
                Dim _OTD As DataSet
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('3.AW SHADE COMMENTS','4.AW CUS APP','5.NEED TO FINISH WITH PENDING APP','6.FINISHED WITH PENDING APP','7.SUBMITTED AS ONGOING','8.AW CUS CARE COMMENTS')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')'   and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('3.AW SHADE COMMENTS','4.AW CUS APP','5.NEED TO FINISH WITH PENDING APP','6.FINISHED WITH PENDING APP','7.SUBMITTED AS ONGOING','8.AW CUS CARE COMMENTS')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')'   and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('3.AW SHADE COMMENTS','4.AW CUS APP','5.NEED TO FINISH WITH PENDING APP','6.FINISHED WITH PENDING APP','7.SUBMITTED AS ONGOING','8.AW CUS CARE COMMENTS')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')'    "

                    End If
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('3.AW SHADE COMMENTS','4.AW CUS APP','5.NEED TO FINISH WITH PENDING APP','6.FINISHED WITH PENDING APP','7.SUBMITTED AS ONGOING','8.AW CUS CARE COMMENTS')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')   and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('3.AW SHADE COMMENTS','4.AW CUS APP','5.NEED TO FINISH WITH PENDING APP','6.FINISHED WITH PENDING APP','7.SUBMITTED AS ONGOING','8.AW CUS CARE COMMENTS')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')    and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('3.AW SHADE COMMENTS','4.AW CUS APP','5.NEED TO FINISH WITH PENDING APP','6.FINISHED WITH PENDING APP','7.SUBMITTED AS ONGOING','8.AW CUS CARE COMMENTS')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')  "

                    End If
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('3.AW SHADE COMMENTS','4.AW CUS APP','5.NEED TO FINISH WITH PENDING APP','6.FINISHED WITH PENDING APP','7.SUBMITTED AS ONGOING','8.AW CUS CARE COMMENTS')  and  Location='Finishing' and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "')'   and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('3.AW SHADE COMMENTS','4.AW CUS APP','5.NEED TO FINISH WITH PENDING APP','6.FINISHED WITH PENDING APP','7.SUBMITTED AS ONGOING','8.AW CUS CARE COMMENTS')  and  Location='Finishing' and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "')'    and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('3.AW SHADE COMMENTS','4.AW CUS APP','5.NEED TO FINISH WITH PENDING APP','6.FINISHED WITH PENDING APP','7.SUBMITTED AS ONGOING','8.AW CUS CARE COMMENTS')  and  Location='Finishing' and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "')'   "

                    End If
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then

                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('3.AW SHADE COMMENTS','4.AW CUS APP','5.NEED TO FINISH WITH PENDING APP','6.FINISHED WITH PENDING APP','7.SUBMITTED AS ONGOING','8.AW CUS CARE COMMENTS')  and  Location='Finishing' and  r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')'   and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('3.AW SHADE COMMENTS','4.AW CUS APP','5.NEED TO FINISH WITH PENDING APP','6.FINISHED WITH PENDING APP','7.SUBMITTED AS ONGOING','8.AW CUS CARE COMMENTS')  and  Location='Finishing'  and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')'    and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('3.AW SHADE COMMENTS','4.AW CUS APP','5.NEED TO FINISH WITH PENDING APP','6.FINISHED WITH PENDING APP','7.SUBMITTED AS ONGOING','8.AW CUS CARE COMMENTS')  and  Location='Finishing'  and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')'   "

                    End If

                ElseIf Trim(txtCustomer.Text) <> "" Then

                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('3.AW SHADE COMMENTS','4.AW CUS APP','5.NEED TO FINISH WITH PENDING APP','6.FINISHED WITH PENDING APP','7.SUBMITTED AS ONGOING','8.AW CUS CARE COMMENTS')  and  Location='Finishing' and s.Customer in ('" & _Customer & "')    and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('3.AW SHADE COMMENTS','4.AW CUS APP','5.NEED TO FINISH WITH PENDING APP','6.FINISHED WITH PENDING APP','7.SUBMITTED AS ONGOING','8.AW CUS CARE COMMENTS')  and  Location='Finishing' and s.Customer in ('" & _Customer & "')   and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('3.AW SHADE COMMENTS','4.AW CUS APP','5.NEED TO FINISH WITH PENDING APP','6.FINISHED WITH PENDING APP','7.SUBMITTED AS ONGOING','8.AW CUS CARE COMMENTS')  and  Location='Finishing' and s.Customer in ('" & _Customer & "')     "

                    End If

                ElseIf Trim(txtDepartment.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('3.AW SHADE COMMENTS','4.AW CUS APP','5.NEED TO FINISH WITH PENDING APP','6.FINISHED WITH PENDING APP','7.SUBMITTED AS ONGOING','8.AW CUS CARE COMMENTS')  and  Location='Finishing' and  r.Department in ('" & _Department & "')   and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('3.AW SHADE COMMENTS','4.AW CUS APP','5.NEED TO FINISH WITH PENDING APP','6.FINISHED WITH PENDING APP','7.SUBMITTED AS ONGOING','8.AW CUS CARE COMMENTS')  and  Location='Finishing' and  r.Department in ('" & _Department & "')  and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('3.AW SHADE COMMENTS','4.AW CUS APP','5.NEED TO FINISH WITH PENDING APP','6.FINISHED WITH PENDING APP','7.SUBMITTED AS ONGOING','8.AW CUS CARE COMMENTS')  and  Location='Finishing' and  r.Department in ('" & _Department & "')     "

                    End If
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('3.AW SHADE COMMENTS','4.AW CUS APP','5.NEED TO FINISH WITH PENDING APP','6.FINISHED WITH PENDING APP','7.SUBMITTED AS ONGOING','8.AW CUS CARE COMMENTS')  and  Location='Finishing' and  r.Merchant in ('" & _Merchant & "')'  and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('3.AW SHADE COMMENTS','4.AW CUS APP','5.NEED TO FINISH WITH PENDING APP','6.FINISHED WITH PENDING APP','7.SUBMITTED AS ONGOING','8.AW CUS CARE COMMENTS')  and  Location='Finishing' and  r.Merchant in ('" & _Merchant & "')'  and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('3.AW SHADE COMMENTS','4.AW CUS APP','5.NEED TO FINISH WITH PENDING APP','6.FINISHED WITH PENDING APP','7.SUBMITTED AS ONGOING','8.AW CUS CARE COMMENTS')  and  Location='Finishing' and  r.Merchant in ('" & _Merchant & "')'  "

                    End If
                Else
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('3.AW SHADE COMMENTS','4.AW CUS APP','5.NEED TO FINISH WITH PENDING APP','6.FINISHED WITH PENDING APP','7.SUBMITTED AS ONGOING','8.AW CUS CARE COMMENTS')  and  Location='Finishing'   and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('3.AW SHADE COMMENTS','4.AW CUS APP','5.NEED TO FINISH WITH PENDING APP','6.FINISHED WITH PENDING APP','7.SUBMITTED AS ONGOING','8.AW CUS CARE COMMENTS')  and  Location='Finishing'  and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('3.AW SHADE COMMENTS','4.AW CUS APP','5.NEED TO FINISH WITH PENDING APP','6.FINISHED WITH PENDING APP','7.SUBMITTED AS ONGOING','8.AW CUS CARE COMMENTS')  and  Location='Finishing'  "

                    End If
                End If

                I = 0
                _QtyKg = 0
                _QtyKg = 0
                _Rowcount = 0
                'T01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetMEETING", New SqlParameter("@cQryType", "ISC"), New SqlParameter("@vcWhereClause", vcWharer))
                'T01 = DBEngin.ExecuteDa
                For Each DTRow3 As DataRow In T01.Tables(0).Rows
                    Dim newRow As DataRow = c_dataCustomer1.NewRow
                    Dim _Tollaranz As Double

                    'For Each DTRow1 As DataRow In M01.Tables(0).Rows


                    Dim _Material As String
                    _Material = T01.Tables(0).Rows(I)("Metrrial")
                    '_Material = Microsoft.VisualBasic.Right(_Material, 7)
                    '_Material = Microsoft.VisualBasic.Left(_Material, 2) & "-" & Microsoft.VisualBasic.Right(_Material, 5)
                    newRow("Sales Order") = T01.Tables(0).Rows(I)("Sales_Order")
                    newRow("Line Item") = T01.Tables(0).Rows(I)("Line_Item")
                    newRow("Material") = _Material
                    newRow("Description") = T01.Tables(0).Rows(I)("Met_Des")
                    newRow("Batch Qty (Mtr)") = Microsoft.VisualBasic.Format(T01.Tables(0).Rows(I)("PRD_Qty"), "#.00")
                    _QtyMtr = _QtyMtr + T01.Tables(0).Rows(I)("PRD_Qty")
                    newRow("Delivary Date") = Month(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Microsoft.VisualBasic.Day(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Year(T01.Tables(0).Rows(I)("Del_Date"))
                    _delivaryDT = Month(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Microsoft.VisualBasic.Day(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Year(T01.Tables(0).Rows(I)("Del_Date"))
                    newRow("Batch No") = T01.Tables(0).Rows(I)("Prduct_Order")
                    newRow("Customer") = T01.Tables(0).Rows(I)("Customer")

                    Sql = "select * from ZPP_DEL where Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and Line_Item='" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & "' and Delivery_Date='" & T01.Tables(0).Rows(I)("Del_Date") & "' and Product_Order='" & Trim(T01.Tables(0).Rows(I)("Prduct_Order")) & "'"
                    dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(dsUser) Then
                        newRow("Batch Qty (Kg)") = Microsoft.VisualBasic.Format(dsUser.Tables(0).Rows(0)("Order_Qty_Kg"), "#.00")
                        _QtyKg = _QtyKg + dsUser.Tables(0).Rows(0)("Order_Qty_Kg")
                        newRow("No.Of.Dys In S/L") = dsUser.Tables(0).Rows(0)("No_Day_Same_Opp")
                        newRow("NC Comment") = dsUser.Tables(0).Rows(0)("NC_Comment")
                        newRow("Next Operation") = dsUser.Tables(0).Rows(0)("Status")
                    End If

                    Sql = "select * from T07OTD_Comment1 where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07Material='" & _Material & "' and T07BatchNo='" & T01.Tables(0).Rows(I)("Prduct_Order") & "' order by T07Date DESC"
                    _OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(_OTD) Then
                        If IsDate(_OTD.Tables(0).Rows(0)("T07Posible_Date")) Then
                            If Year(_OTD.Tables(0).Rows(0)("T07Posible_Date")) = "1900" Then
                            Else
                                newRow("Possible Devivary Date") = _OTD.Tables(0).Rows(0)("T07Posible_Date")

                            End If
                        End If

                        newRow("Reason") = _OTD.Tables(0).Rows(0)("T07Reason")
                        If _OTD.Tables(0).Rows(0)("T07date") = Today Then
                            newRow(_Todayupdate) = _OTD.Tables(0).Rows(0)("T07Comment")
                        Else
                            newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")
                        End If
                        newRow("LIB dep") = _OTD.Tables(0).Rows(0)("T07LIB")

                        _YesDate = Today.AddDays(-1)
                        Sql = "select * from T07OTD_Comment1 where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07BatchNo='" & T01.Tables(0).Rows(I)("Prduct_Order") & "' and T07Date<'" & Today & "' order by T07Date DESC"
                        _OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                        If isValidDataset(_OTD) Then
                            newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")
                        End If

                    Else
                        'Sql = "select * from T07OTD_Comment where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07Material='" & _Material & "'  order by T07Date DESC"
                        '_OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                        'If isValidDataset(_OTD) Then
                        '    ' newRow("Possible Devivary Date") = _OTD.Tables(0).Rows(0)("T07Posible_Date")
                        '    ' newRow("Reason") = _OTD.Tables(0).Rows(0)("T07Reason")
                        '    newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")
                        '    ' newRow("LIB dep") = _OTD.Tables(0).Rows(0)("T07LIB")
                        'End If
                    End If
                    newRow("Week") = "Week " & DatePart(DateInterval.WeekOfYear, _delivaryDT)
                    c_dataCustomer1.Rows.Add(newRow)
                    ' End If

                    ' If Not x = M01.Tables(0).Rows.Count - 1 Then
                    ' rsT24RecHeader.MoveNext()
                    _Rowcount = _Rowcount + 1
                    I = I + 1
                Next
            ElseIf Trim(cboPO.Text) = "Aw  customer Shade comment" Then
                Dim _OTD As DataSet
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('2.2 AW N/C APP','2.3 AW CUS APP','2.7 SUBMIT AS ONGOING')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('2.2 AW N/C APP','2.3 AW CUS APP','2.7 SUBMIT AS ONGOING')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('2.2 AW N/C APP','2.3 AW CUS APP','2.7 SUBMIT AS ONGOING')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') "

                    End If
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('2.2 AW N/C APP','2.3 AW CUS APP','2.7 SUBMIT AS ONGOING')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')  and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('2.2 AW N/C APP','2.3 AW CUS APP','2.7 SUBMIT AS ONGOING')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')  and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('2.2 AW N/C APP','2.3 AW CUS APP','2.7 SUBMIT AS ONGOING')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')  "

                    End If
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('2.2 AW N/C APP','2.3 AW CUS APP','2.7 SUBMIT AS ONGOING')  and  Location='Finishing' and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "') and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('2.2 AW N/C APP','2.3 AW CUS APP','2.7 SUBMIT AS ONGOING')  and  Location='Finishing' and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "') and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('2.2 AW N/C APP','2.3 AW CUS APP','2.7 SUBMIT AS ONGOING')  and  Location='Finishing' and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "') "

                    End If
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then

                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('2.2 AW N/C APP','2.3 AW CUS APP','2.7 SUBMIT AS ONGOING')  and  Location='Finishing'  and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('2.2 AW N/C APP','2.3 AW CUS APP','2.7 SUBMIT AS ONGOING')  and  Location='Finishing'  and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('2.2 AW N/C APP','2.3 AW CUS APP','2.7 SUBMIT AS ONGOING')  and  Location='Finishing'  and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') "

                    End If

                ElseIf Trim(txtCustomer.Text) <> "" Then

                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('2.2 AW N/C APP','2.3 AW CUS APP','2.7 SUBMIT AS ONGOING')  and  Location='Finishing' and s.Customer in ('" & _Customer & "')  and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('2.2 AW N/C APP','2.3 AW CUS APP','2.7 SUBMIT AS ONGOING')  and  Location='Finishing' and s.Customer in ('" & _Customer & "')  and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('2.2 AW N/C APP','2.3 AW CUS APP','2.7 SUBMIT AS ONGOING')  and  Location='Finishing' and s.Customer in ('" & _Customer & "')  "

                    End If

                ElseIf Trim(txtDepartment.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('2.2 AW N/C APP','2.3 AW CUS APP','2.7 SUBMIT AS ONGOING')  and  Location='Finishing' and r.Department in ('" & _Department & "')  and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('2.2 AW N/C APP','2.3 AW CUS APP','2.7 SUBMIT AS ONGOING')  and  Location='Finishing' and r.Department in ('" & _Department & "') and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('2.2 AW N/C APP','2.3 AW CUS APP','2.7 SUBMIT AS ONGOING')  and  Location='Finishing' and  r.Department in ('" & _Department & "')  "

                    End If
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('2.2 AW N/C APP','2.3 AW CUS APP','2.7 SUBMIT AS ONGOING')  and  Location='Finishing'  and r.Merchant in ('" & _Merchant & "') and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('2.2 AW N/C APP','2.3 AW CUS APP','2.7 SUBMIT AS ONGOING')  and  Location='Finishing'  and r.Merchant in ('" & _Merchant & "') and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('2.2 AW N/C APP','2.3 AW CUS APP','2.7 SUBMIT AS ONGOING')  and  Location='Finishing'  and r.Merchant in ('" & _Merchant & "') "

                    End If
                Else
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('2.2 AW N/C APP','2.3 AW CUS APP','2.7 SUBMIT AS ONGOING')  and  Location='Finishing'  and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('2.2 AW N/C APP','2.3 AW CUS APP','2.7 SUBMIT AS ONGOING')  and  Location='Finishing'  and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('2.2 AW N/C APP','2.3 AW CUS APP','2.7 SUBMIT AS ONGOING')  and  Location='Finishing'  "

                    End If
                End If

                I = 0
                'T01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetMEETING", New SqlParameter("@cQryType", "ICS"), New SqlParameter("@vcWhereClause", vcWharer))
                'T01 = DBEngin.ExecuteDa
                _QtyKg = 0
                _QtyMtr = 0
                _Rowcount = 0

                For Each DTRow3 As DataRow In T01.Tables(0).Rows
                    Dim newRow As DataRow = c_dataCustomer1.NewRow
                    Dim _Tollaranz As Double

                    'For Each DTRow1 As DataRow In M01.Tables(0).Rows


                    Dim _Material As String
                    _Material = T01.Tables(0).Rows(I)("Metrrial")
                    '_Material = Microsoft.VisualBasic.Right(_Material, 7)
                    '_Material = Microsoft.VisualBasic.Left(_Material, 2) & "-" & Microsoft.VisualBasic.Right(_Material, 5)
                    newRow("Sales Order") = T01.Tables(0).Rows(I)("Sales_Order")
                    newRow("Line Item") = T01.Tables(0).Rows(I)("Line_Item")
                    newRow("Material") = _Material
                    newRow("Description") = T01.Tables(0).Rows(I)("Met_Des")
                    newRow("Batch Qty (Mtr)") = Microsoft.VisualBasic.Format(T01.Tables(0).Rows(I)("PRD_Qty"), "#.00")
                    _QtyMtr = _QtyMtr + T01.Tables(0).Rows(I)("PRD_Qty")
                    newRow("Delivary Date") = Month(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Microsoft.VisualBasic.Day(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Year(T01.Tables(0).Rows(I)("Del_Date"))
                    _delivaryDT = Month(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Microsoft.VisualBasic.Day(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Year(T01.Tables(0).Rows(I)("Del_Date"))

                    newRow("Batch No") = T01.Tables(0).Rows(I)("Prduct_Order")
                    newRow("Customer") = T01.Tables(0).Rows(I)("Customer")

                    Sql = "select * from ZPP_DEL where Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and Line_Item='" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & "' and Delivery_Date='" & T01.Tables(0).Rows(I)("Del_Date") & "' and Product_Order='" & Trim(T01.Tables(0).Rows(I)("Prduct_Order")) & "'"
                    dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(dsUser) Then
                        newRow("Batch Qty (Kg)") = Microsoft.VisualBasic.Format(dsUser.Tables(0).Rows(0)("Order_Qty_Kg"), "#.00")
                        _QtyKg = _QtyKg + dsUser.Tables(0).Rows(0)("Order_Qty_Kg")
                        newRow("No.Of.Dys In S/L") = dsUser.Tables(0).Rows(0)("No_Day_Same_Opp")
                        newRow("NC Comment") = dsUser.Tables(0).Rows(0)("NC_Comment")
                        newRow("Next Operation") = dsUser.Tables(0).Rows(0)("Status")
                    End If

                    Sql = "select * from T07OTD_Comment1 where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07Material='" & _Material & "' and T07BatchNo='" & T01.Tables(0).Rows(I)("Prduct_Order") & "' order by T07Date DESC"
                    _OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(_OTD) Then
                        If IsDate(_OTD.Tables(0).Rows(0)("T07Posible_Date")) Then
                            If Year(_OTD.Tables(0).Rows(0)("T07Posible_Date")) = "1900" Then
                            Else
                                newRow("Possible Devivary Date") = _OTD.Tables(0).Rows(0)("T07Posible_Date")

                            End If
                        End If

                        newRow("Reason") = _OTD.Tables(0).Rows(0)("T07Reason")
                        If _OTD.Tables(0).Rows(0)("T07date") = Today Then
                            newRow(_Todayupdate) = _OTD.Tables(0).Rows(0)("T07Comment")
                        Else
                            newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")
                        End If
                        newRow("LIB dep") = _OTD.Tables(0).Rows(0)("T07LIB")

                        _YesDate = Today.AddDays(-1)
                        Sql = "select * from T07OTD_Comment1 where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07BatchNo='" & T01.Tables(0).Rows(I)("Prduct_Order") & "' and T07Date<'" & Today & "' order by T07Date DESC"
                        _OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                        If isValidDataset(_OTD) Then
                            newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")
                        End If

                    Else
                        'Sql = "select * from T07OTD_Comment where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07Material='" & _Material & "'  order by T07Date DESC"
                        '_OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                        'If isValidDataset(_OTD) Then
                        '    'newRow("Possible Devivary Date") = _OTD.Tables(0).Rows(0)("T07Posible_Date")
                        '    'newRow("Reason") = _OTD.Tables(0).Rows(0)("T07Reason")
                        '    newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")
                        '    'newRow("LIB dep") = _OTD.Tables(0).Rows(0)("T07LIB")
                        'End If
                    End If

                    newRow("Week") = "Week " & DatePart(DateInterval.WeekOfYear, _delivaryDT)
                    c_dataCustomer1.Rows.Add(newRow)
                    ' End If

                    ' If Not x = M01.Tables(0).Rows.Count - 1 Then
                    ' rsT24RecHeader.MoveNext()
                    _Rowcount = _Rowcount + 1
                    I = I + 1
                Next

            ElseIf Trim(cboPO.Text) = "Held in  N/C due to  dyeing issues" Then
                Dim _OTD As DataSet
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('11.Held in other reason dyeing','17.Held due to Trials','28.DOWN  GRADE')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('11.Held in other reason dyeing','17.Held due to Trials','28.DOWN  GRADE')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('11.Held in other reason dyeing','17.Held due to Trials','28.DOWN  GRADE')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  "

                    End If
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('11.Held in other reason dyeing','17.Held due to Trials','28.DOWN  GRADE')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')  and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('11.Held in other reason dyeing','17.Held due to Trials','28.DOWN  GRADE')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')   and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('11.Held in other reason dyeing','17.Held due to Trials','28.DOWN  GRADE')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')   "

                    End If
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('11.Held in other reason dyeing','17.Held due to Trials','28.DOWN  GRADE')  and  Location='Finishing' and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "')  and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('11.Held in other reason dyeing','17.Held due to Trials','28.DOWN  GRADE')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Merchant in ('" & _Merchant & "')  and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('11.Held in other reason dyeing','17.Held due to Trials','28.DOWN  GRADE')  and  Location='Finishing' and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "')  "

                    End If
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then

                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('11.Held in other reason dyeing','17.Held due to Trials','28.DOWN  GRADE')  and  Location='Finishing'  and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('11.Held in other reason dyeing','17.Held due to Trials','28.DOWN  GRADE')  and  Location='Finishing'  and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('11.Held in other reason dyeing','17.Held due to Trials','28.DOWN  GRADE') and  Location='Finishing'  and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  "

                    End If

                ElseIf Trim(txtCustomer.Text) <> "" Then

                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('11.Held in other reason dyeing','17.Held due to Trials','28.DOWN  GRADE')  and  Location='Finishing' and s.Customer in ('" & _Customer & "')   and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('11.Held in other reason dyeing','17.Held due to Trials','28.DOWN  GRADE')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Merchant in ('" & _Merchant & "')  and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('11.Held in other reason dyeing','17.Held due to Trials','28.DOWN  GRADE')  and  Location='Finishing' and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "')  "

                    End If

                ElseIf Trim(txtDepartment.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('11.Held in other reason dyeing','17.Held due to Trials','28.DOWN  GRADE')  and  Location='Finishing'  and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('11.Held in other reason dyeing','17.Held due to Trials','28.DOWN  GRADE')  and  Location='Finishing' and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('11.Held in other reason dyeing','17.Held due to Trials','28.DOWN  GRADE')  and  Location='Finishing'  and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  "

                    End If
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('11.Held in other reason dyeing','17.Held due to Trials','28.DOWN  GRADE')  and  Location='Finishing'  and r.Merchant in ('" & _Merchant & "')  and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('11.Held in other reason dyeing','17.Held due to Trials','28.DOWN  GRADE')  and  Location='Finishing'  and r.Merchant in ('" & _Merchant & "')  and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('11.Held in other reason dyeing','17.Held due to Trials','28.DOWN  GRADE') and  Location='Finishing'  and r.Merchant in ('" & _Merchant & "')  "

                    End If
                Else
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('11.Held in other reason dyeing','17.Held due to Trials','28.DOWN  GRADE')  and  Location='Finishing'   and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('11.Held in other reason dyeing','17.Held due to Trials','28.DOWN  GRADE')  and  Location='Finishing'   and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('11.Held in other reason dyeing','17.Held due to Trials','28.DOWN  GRADE')  and  Location='Finishing'   "

                    End If
                End If

                _Rowcount = 0
                _QtyKg = 0
                _QtyMtr = 0
                I = 0
                'T01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetMEETING", New SqlParameter("@cQryType", "GRG"), New SqlParameter("@vcWhereClause", vcWharer))
                'T01 = DBEngin.ExecuteDa
                For Each DTRow3 As DataRow In T01.Tables(0).Rows
                    Dim newRow As DataRow = c_dataCustomer1.NewRow
                    Dim _Tollaranz As Double

                    'For Each DTRow1 As DataRow In M01.Tables(0).Rows


                    Dim _Material As String
                    _Material = T01.Tables(0).Rows(I)("Metrrial")
                    '_Material = Microsoft.VisualBasic.Right(_Material, 7)
                    '_Material = Microsoft.VisualBasic.Left(_Material, 2) & "-" & Microsoft.VisualBasic.Right(_Material, 5)
                    newRow("Sales Order") = T01.Tables(0).Rows(I)("Sales_Order")
                    newRow("Line Item") = T01.Tables(0).Rows(I)("Line_Item")
                    newRow("Material") = _Material
                    newRow("Description") = T01.Tables(0).Rows(I)("Met_Des")
                    newRow("Batch Qty (Mtr)") = Microsoft.VisualBasic.Format(T01.Tables(0).Rows(I)("PRD_Qty"), "#.00")
                    _QtyMtr = _QtyMtr + T01.Tables(0).Rows(I)("PRD_Qty")
                    newRow("Delivary Date") = Month(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Microsoft.VisualBasic.Day(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Year(T01.Tables(0).Rows(I)("Del_Date"))
                    _delivaryDT = Month(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Microsoft.VisualBasic.Day(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Year(T01.Tables(0).Rows(I)("Del_Date"))

                    newRow("Batch No") = T01.Tables(0).Rows(I)("Prduct_Order")
                    newRow("Customer") = T01.Tables(0).Rows(I)("Customer")

                    Sql = "select * from ZPP_DEL where Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and Line_Item='" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & "' and Delivery_Date='" & T01.Tables(0).Rows(I)("Del_Date") & "' and Product_Order='" & Trim(T01.Tables(0).Rows(I)("Prduct_Order")) & "'"
                    dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(dsUser) Then
                        newRow("Batch Qty (Kg)") = Microsoft.VisualBasic.Format(dsUser.Tables(0).Rows(0)("Order_Qty_Kg"), "#.00")
                        _QtyKg = _QtyKg + dsUser.Tables(0).Rows(0)("Order_Qty_Kg")
                        newRow("No.Of.Dys In S/L") = dsUser.Tables(0).Rows(0)("No_Day_Same_Opp")
                        newRow("NC Comment") = dsUser.Tables(0).Rows(0)("NC_Comment")
                        newRow("Next Operation") = dsUser.Tables(0).Rows(0)("Status")
                    End If

                    Sql = "select * from T07OTD_Comment1 where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07Material='" & _Material & "' and T07BatchNo='" & T01.Tables(0).Rows(I)("Prduct_Order") & "' order by T07Date DESC"
                    _OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(_OTD) Then
                        If IsDate(_OTD.Tables(0).Rows(0)("T07Posible_Date")) Then
                            If Year(_OTD.Tables(0).Rows(0)("T07Posible_Date")) = "1900" Then
                            Else
                                newRow("Possible Devivary Date") = _OTD.Tables(0).Rows(0)("T07Posible_Date")

                            End If
                        End If

                        newRow("Reason") = _OTD.Tables(0).Rows(0)("T07Reason")
                        If _OTD.Tables(0).Rows(0)("T07date") = Today Then
                            newRow(_Todayupdate) = _OTD.Tables(0).Rows(0)("T07Comment")
                        Else
                            newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")
                        End If
                        newRow("LIB dep") = _OTD.Tables(0).Rows(0)("T07LIB")

                        _YesDate = Today.AddDays(-1)
                        Sql = "select * from T07OTD_Comment1 where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07BatchNo='" & T01.Tables(0).Rows(I)("Prduct_Order") & "' and T07Date<'" & Today & "' order by T07Date DESC"
                        _OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                        If isValidDataset(_OTD) Then
                            newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")
                        End If

                    Else
                        'Sql = "select * from T07OTD_Comment where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07Material='" & _Material & "'  order by T07Date DESC"
                        '_OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                        'If isValidDataset(_OTD) Then
                        '    'newRow("Possible Devivary Date") = _OTD.Tables(0).Rows(0)("T07Posible_Date")
                        '    'newRow("Reason") = _OTD.Tables(0).Rows(0)("T07Reason")
                        '    newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")
                        '    'newRow("LIB dep") = _OTD.Tables(0).Rows(0)("T07LIB")
                        'End If
                    End If
                    newRow("Week") = "Week " & DatePart(DateInterval.WeekOfYear, _delivaryDT)
                    c_dataCustomer1.Rows.Add(newRow)
                    ' End If

                    ' If Not x = M01.Tables(0).Rows.Count - 1 Then
                    ' rsT24RecHeader.MoveNext()
                    _Rowcount = _Rowcount + 1
                    I = I + 1
                Next

            ElseIf Trim(cboPO.Text) = "Held in  N/C due to  Finishing issues" Then
                Dim _OTD As DataSet
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('4.2 FINISHING ISSUES')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')   and left(r.Met_Des,8)='" & cboQuality.Text & "' and r.NC_Comment in ('12.Held in other reason Finishing','14.HELD IN OTHER REASON  OTHERS')"
                        Else
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('4.2 FINISHING ISSUES')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')   and left(r.Met_Des,5)='" & cboQuality.Text & "' and r.NC_Comment in ('12.Held in other reason Finishing','14.HELD IN OTHER REASON  OTHERS')"
                        End If
                    Else
                        vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('4.2 FINISHING ISSUES')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.NC_Comment in ('12.Held in other reason Finishing','14.HELD IN OTHER REASON  OTHERS') "

                    End If
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('4.2 FINISHING ISSUES')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')    and left(r.Met_Des,8)='" & cboQuality.Text & "' and r.NC_Comment in ('12.Held in other reason Finishing','14.HELD IN OTHER REASON  OTHERS')"
                        Else
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('4.2 FINISHING ISSUES')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')   and left(r.Met_Des,5)='" & cboQuality.Text & "' and r.NC_Comment in ('12.Held in other reason Finishing','14.HELD IN OTHER REASON  OTHERS')"
                        End If
                    Else
                        vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('4.2 FINISHING ISSUES')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')  "

                    End If
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('4.2 FINISHING ISSUES')  and  Location='Finishing' and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "')   and left(r.Met_Des,8)='" & cboQuality.Text & "' and r.NC_Comment in ('12.Held in other reason Finishing','14.HELD IN OTHER REASON  OTHERS')"
                        Else
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('4.2 FINISHING ISSUES')  and  Location='Finishing' and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "')   and left(r.Met_Des,5)='" & cboQuality.Text & "' and r.NC_Comment in ('12.Held in other reason Finishing','14.HELD IN OTHER REASON  OTHERS')"
                        End If
                    Else
                        vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('4.2 FINISHING ISSUES')  and  Location='Finishing' and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "') and r.NC_Comment in ('12.Held in other reason Finishing','14.HELD IN OTHER REASON  OTHERS') "

                    End If
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then

                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('4.2 FINISHING ISSUES')  and  Location='Finishing' and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')   and left(r.Met_Des,8)='" & cboQuality.Text & "' and r.NC_Comment in ('12.Held in other reason Finishing','14.HELD IN OTHER REASON  OTHERS')"
                        Else
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('4.2 FINISHING ISSUES')  and  Location='Finishing' and  and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')   and left(r.Met_Des,5)='" & cboQuality.Text & "' and r.NC_Comment in ('12.Held in other reason Finishing','14.HELD IN OTHER REASON  OTHERS')"
                        End If
                    Else
                        vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('4.2 FINISHING ISSUES')  and  Location='Finishing' and  r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.NC_Comment in ('12.Held in other reason Finishing','14.HELD IN OTHER REASON  OTHERS') "

                    End If

                ElseIf Trim(txtCustomer.Text) <> "" Then

                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('4.2 FINISHING ISSUES')  and  Location='Finishing' and s.Customer in ('" & _Customer & "')   and left(r.Met_Des,8)='" & cboQuality.Text & "' and r.NC_Comment in ('12.Held in other reason Finishing','14.HELD IN OTHER REASON  OTHERS')"
                        Else
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('4.2 FINISHING ISSUES')  and  Location='Finishing' and s.Customer in ('" & _Customer & "')   and left(r.Met_Des,5)='" & cboQuality.Text & "' and r.NC_Comment in ('12.Held in other reason Finishing','14.HELD IN OTHER REASON  OTHERS')"
                        End If
                    Else
                        vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('4.2 FINISHING ISSUES')  and  Location='Finishing' and s.Customer in ('" & _Customer & "')  and r.NC_Comment in ('12.Held in other reason Finishing','14.HELD IN OTHER REASON  OTHERS') "

                    End If

                ElseIf Trim(txtDepartment.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('4.2 FINISHING ISSUES')  and  Location='Finishing'  and r.Department in ('" & _Department & "')   and left(r.Met_Des,8)='" & cboQuality.Text & "' and r.NC_Comment in ('12.Held in other reason Finishing','14.HELD IN OTHER REASON  OTHERS')"
                        Else
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('4.2 FINISHING ISSUES')  and  Location='Finishing'  and r.Department in ('" & _Department & "')    and left(r.Met_Des,5)='" & cboQuality.Text & "' and r.NC_Comment in ('12.Held in other reason Finishing','14.HELD IN OTHER REASON  OTHERS')"
                        End If
                    Else
                        vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('4.2 FINISHING ISSUES')  and  Location='Finishing'  and r.Department in ('" & _Department & "') and r.NC_Comment in ('12.Held in other reason Finishing','14.HELD IN OTHER REASON  OTHERS')  "

                    End If
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('4.2 FINISHING ISSUES')  and  Location='Finishing'  and r.Merchant in ('" & _Merchant & "')   and left(r.Met_Des,8)='" & cboQuality.Text & "' and r.NC_Comment in ('12.Held in other reason Finishing','14.HELD IN OTHER REASON  OTHERS')"
                        Else
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('4.2 FINISHING ISSUES')  and  Location='Finishing'  and r.Merchant in ('" & _Merchant & "')   and left(r.Met_Des,5)='" & cboQuality.Text & "' and r.NC_Comment in ('12.Held in other reason Finishing','14.HELD IN OTHER REASON  OTHERS')"
                        End If
                    Else
                        vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('4.2 FINISHING ISSUES')  and  Location='Finishing'  and r.Merchant in ('" & _Merchant & "') and r.NC_Comment in ('12.Held in other reason Finishing','14.HELD IN OTHER REASON  OTHERS') "

                    End If
                Else
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('4.2 FINISHING ISSUES')  and  Location='Finishing'    and left(r.Met_Des,8)='" & cboQuality.Text & "' and r.NC_Comment in ('12.Held in other reason Finishing','14.HELD IN OTHER REASON  OTHERS')"
                        Else
                            vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('4.2 FINISHING ISSUES')  and  Location='Finishing'    and left(r.Met_Des,5)='" & cboQuality.Text & "' and r.NC_Comment in ('12.Held in other reason Finishing','14.HELD IN OTHER REASON  OTHERS')"
                        End If
                    Else
                        vcWharer = "  s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('4.2 FINISHING ISSUES')  and  Location='Finishing' and r.NC_Comment in ('12.Held in other reason Finishing','14.HELD IN OTHER REASON  OTHERS')  "

                    End If
                End If

                I = 0
                'T01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetMEETING", New SqlParameter("@cQryType", "GRG"), New SqlParameter("@vcWhereClause", vcWharer))
                'T01 = DBEngin.ExecuteDa
                _Rowcount = 0
                _QtyKg = 0
                _QtyMtr = 0
                For Each DTRow3 As DataRow In T01.Tables(0).Rows
                    Dim newRow As DataRow = c_dataCustomer1.NewRow
                    Dim _Tollaranz As Double

                    'For Each DTRow1 As DataRow In M01.Tables(0).Rows


                    Dim _Material As String
                    _Material = T01.Tables(0).Rows(I)("Metrrial")
                    '_Material = Microsoft.VisualBasic.Right(_Material, 7)
                    '_Material = Microsoft.VisualBasic.Left(_Material, 2) & "-" & Microsoft.VisualBasic.Right(_Material, 5)
                    newRow("Sales Order") = T01.Tables(0).Rows(I)("Sales_Order")
                    newRow("Line Item") = T01.Tables(0).Rows(I)("Line_Item")
                    newRow("Material") = _Material
                    newRow("Description") = T01.Tables(0).Rows(I)("Met_Des")
                    newRow("Batch Qty (Mtr)") = Microsoft.VisualBasic.Format(T01.Tables(0).Rows(I)("PRD_Qty"), "#.00")
                    _QtyMtr = _QtyMtr + T01.Tables(0).Rows(I)("PRD_Qty")
                    newRow("Delivary Date") = Month(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Microsoft.VisualBasic.Day(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Year(T01.Tables(0).Rows(I)("Del_Date"))
                    _delivaryDT = Month(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Microsoft.VisualBasic.Day(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Year(T01.Tables(0).Rows(I)("Del_Date"))
                    newRow("Batch No") = T01.Tables(0).Rows(I)("Prduct_Order")
                    newRow("Customer") = T01.Tables(0).Rows(I)("Customer")

                    Sql = "select * from ZPP_DEL where Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and Line_Item='" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & "' and Delivery_Date='" & T01.Tables(0).Rows(I)("Del_Date") & "' and Product_Order='" & Trim(T01.Tables(0).Rows(I)("Prduct_Order")) & "'"
                    dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(dsUser) Then
                        newRow("Batch Qty (Kg)") = Microsoft.VisualBasic.Format(dsUser.Tables(0).Rows(0)("Order_Qty_Kg"), "#.00")
                        _QtyKg = _QtyKg + dsUser.Tables(0).Rows(0)("Order_Qty_Kg")
                        newRow("No.Of.Dys In S/L") = dsUser.Tables(0).Rows(0)("No_Day_Same_Opp")
                        newRow("NC Comment") = dsUser.Tables(0).Rows(0)("NC_Comment")
                        newRow("Next Operation") = dsUser.Tables(0).Rows(0)("Status")
                    End If

                    Sql = "select * from T07OTD_Comment1 where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07Material='" & _Material & "' and T07BatchNo='" & T01.Tables(0).Rows(I)("Prduct_Order") & "' order by T07Date DESC"
                    _OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(_OTD) Then
                        If IsDate(_OTD.Tables(0).Rows(0)("T07Posible_Date")) Then
                            If Year(_OTD.Tables(0).Rows(0)("T07Posible_Date")) = "1900" Then
                            Else
                                newRow("Possible Devivary Date") = _OTD.Tables(0).Rows(0)("T07Posible_Date")

                            End If
                        End If

                        newRow("Reason") = _OTD.Tables(0).Rows(0)("T07Reason")
                        If _OTD.Tables(0).Rows(0)("T07date") = Today Then
                            newRow(_Todayupdate) = _OTD.Tables(0).Rows(0)("T07Comment")
                        Else
                            newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")
                        End If
                        newRow("LIB dep") = _OTD.Tables(0).Rows(0)("T07LIB")

                        _YesDate = Today.AddDays(-1)
                        Sql = "select * from T07OTD_Comment1 where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07BatchNo='" & T01.Tables(0).Rows(I)("Prduct_Order") & "' and T07Date<'" & Today & "' order by T07Date DESC"
                        _OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                        If isValidDataset(_OTD) Then
                            newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")
                        End If


                    Else
                        'Sql = "select * from T07OTD_Comment where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07Material='" & _Material & "'  order by T07Date DESC"
                        '_OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                        'If isValidDataset(_OTD) Then
                        '    'newRow("Possible Devivary Date") = _OTD.Tables(0).Rows(0)("T07Posible_Date")
                        '    'newRow("Reason") = _OTD.Tables(0).Rows(0)("T07Reason")
                        '    newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")
                        '    'newRow("LIB dep") = _OTD.Tables(0).Rows(0)("T07LIB")
                        'End If
                    End If
                    newRow("Week") = "Week " & DatePart(DateInterval.WeekOfYear, _delivaryDT)
                    c_dataCustomer1.Rows.Add(newRow)
                    ' End If
                    _Rowcount = _Rowcount + 1
                    ' If Not x = M01.Tables(0).Rows.Count - 1 Then
                    ' rsT24RecHeader.MoveNext()

                    I = I + 1
                Next

            ElseIf Trim(cboPO.Text) = "Aw Finishing" Then

                Dim _OTD As DataSet
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'   and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and NC_Comment in ('1.Aw 1st bulk Pilot','2.AW ONGOING PILOT','29.AW PICKING','') and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'   and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and NC_Comment in ('1.Aw 1st bulk Pilot','2.AW ONGOING PILOT','29.AW PICKING','') and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'   and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and NC_Comment in ('1.Aw 1st bulk Pilot','2.AW ONGOING PILOT','29.AW PICKING','') "

                    End If
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'   and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and NC_Comment in ('1.Aw 1st bulk Pilot','2.AW ONGOING PILOT','29.AW PICKING','') and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'   and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')  and NC_Comment in ('1.Aw 1st bulk Pilot','2.AW ONGOING PILOT','29.AW PICKING','') and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'   and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')  and NC_Comment in ('1.Aw 1st bulk Pilot','2.AW ONGOING PILOT','29.AW PICKING','') "

                    End If
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'   and  Location='Finishing' and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "') and NC_Comment in ('1.Aw 1st bulk Pilot','2.AW ONGOING PILOT','29.AW PICKING','') and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'   and  Location='Finishing' and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "') and NC_Comment in ('1.Aw 1st bulk Pilot','2.AW ONGOING PILOT','29.AW PICKING','') and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'   and  Location='Finishing' and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "') and NC_Comment in ('1.Aw 1st bulk Pilot','2.AW ONGOING PILOT','29.AW PICKING','') "

                    End If
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then

                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'   and  Location='Finishing' and  r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and NC_Comment in ('1.Aw 1st bulk Pilot','2.AW ONGOING PILOT','29.AW PICKING','') and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'   and  Location='Finishing' and  r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and NC_Comment in ('1.Aw 1st bulk Pilot','2.AW ONGOING PILOT','29.AW PICKING','') and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'   and  Location='Finishing' and  r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and NC_Comment in ('1.Aw 1st bulk Pilot','2.AW ONGOING PILOT','29.AW PICKING','') "

                    End If

                ElseIf Trim(txtCustomer.Text) <> "" Then

                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'   and  Location='Finishing' and s.Customer in ('" & _Customer & "')  and NC_Comment in ('1.Aw 1st bulk Pilot','2.AW ONGOING PILOT','29.AW PICKING','') and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'   and  Location='Finishing' and s.Customer in ('" & _Customer & "')  and NC_Comment in ('1.Aw 1st bulk Pilot','2.AW ONGOING PILOT','29.AW PICKING','') and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'   and  Location='Finishing' and s.Customer in ('" & _Customer & "')  and NC_Comment in ('1.Aw 1st bulk Pilot','2.AW ONGOING PILOT','29.AW PICKING','') "

                    End If

                ElseIf Trim(txtDepartment.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'   and  Location='Finishing'  and r.Department in ('" & _Department & "')  and NC_Comment in ('1.Aw 1st bulk Pilot','2.AW ONGOING PILOT','29.AW PICKING','') and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'   and  Location='Finishing'  and r.Department in ('" & _Department & "')  and NC_Comment in ('1.Aw 1st bulk Pilot','2.AW ONGOING PILOT','29.AW PICKING','') and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'   and  Location='Finishing' and r.Department in ('" & _Department & "')  and NC_Comment in ('1.Aw 1st bulk Pilot','2.AW ONGOING PILOT','29.AW PICKING','') "

                    End If
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'   and  Location='Finishing'  and r.Merchant in ('" & _Merchant & "') and NC_Comment in ('1.Aw 1st bulk Pilot','2.AW ONGOING PILOT','29.AW PICKING','') and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'   and  Location='Finishing'  and r.Merchant in ('" & _Merchant & "') and NC_Comment in ('1.Aw 1st bulk Pilot','2.AW ONGOING PILOT','29.AW PICKING','') and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'   and  Location='Finishing'  and r.Merchant in ('" & _Merchant & "') and NC_Comment in ('1.Aw 1st bulk Pilot','2.AW ONGOING PILOT','29.AW PICKING','') "

                    End If
                Else
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'   and  Location='Finishing'  and NC_Comment in ('1.Aw 1st bulk Pilot','2.AW ONGOING PILOT','29.AW PICKING','') and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'   and  Location='Finishing'  and NC_Comment in ('1.Aw 1st bulk Pilot','2.AW ONGOING PILOT','29.AW PICKING','') and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'   and  Location='Finishing'  and NC_Comment in ('1.Aw 1st bulk Pilot','2.AW ONGOING PILOT','29.AW PICKING','') "

                    End If
                End If

                I = 0
                'T01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetMEETING", New SqlParameter("@cQryType", "ICS"), New SqlParameter("@vcWhereClause", vcWharer))
                'T01 = DBEngin.ExecuteDa
                _Rowcount = 0
                _QtyKg = 0
                _QtyMtr = 0
                For Each DTRow3 As DataRow In T01.Tables(0).Rows
                    Dim newRow As DataRow = c_dataCustomer1.NewRow
                    Dim _Tollaranz As Double
                    'Dim T03 As DataSet

                    'For Each DTRow1 As DataRow In M01.Tables(0).Rows

                    'Sql = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,f.Stock_Code,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date inner join FR_Update F on f.Batch_No=r.Prduct_Order where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location IN ('Finishing') and f.Recipy_Status in ('Awaiting Recipe','First Bulk Awaiting Recipe')  and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') OR R.NC_Comment IN ('4.5 DRY AND HOLD','3.3 OFF SHADE','3.4 NEED TO OVER DYE','3.0 Shade Issues') AND R.Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and r.Line_Item='" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & "'"
                    ''COMMENT ON 2015.4.5
                    ''Sql = "select R.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,f.Stock_Code,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date inner join FR_Update F on f.Batch_No=r.Prduct_Order where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location IN ('Finishing') and f.Recipy_Status in ('Awaiting Recipe','First Bulk Awaiting Recipe')  and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  AND R.Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and r.Line_Item='" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & "'"
                    'T03 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    'If isValidDataset(T03) Then

                    'Else

                    '    Sql = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('00 PILOT','2.1 AW 1ST BULK APP','2.6 AW CUS CARE COMMENT')  and  Location='Finishing' AND R.Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and r.Line_Item='" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & "'"
                    '    T03 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    '    If isValidDataset(T03) Then
                    '    Else
                    '        Sql = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('4.0 HOLD','3.5 OTHERS','3.0 SHADE ISSUES','4.0 OTHER REASON','6.0 HELD IN NEW ORDERS','4.1 DYEING ISSUES','')  and  Location='Finishing' AND R.Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' "
                    '        T03 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    '        If isValidDataset(T03) Then
                    '        Else
                    '            Sql = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('4.2 FINISHING ISSUES','5.2 NEED TO WASH','4.5 DRY AND HOLD')  and  Location='Finishing' AND R.Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and r.Line_Item='" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & "'"
                    '            T03 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    '            If isValidDataset(T03) Then
                    '            Else
                    '                Dim _Material As String
                    '                _Material = T01.Tables(0).Rows(I)("Metrrial")
                    '                '_Material = Microsoft.VisualBasic.Right(_Material, 7)
                    '                '_Material = Microsoft.VisualBasic.Left(_Material, 2) & "-" & Microsoft.VisualBasic.Right(_Material, 5)
                    '                newRow("Sales Order") = T01.Tables(0).Rows(I)("Sales_Order")
                    '                newRow("Line Item") = T01.Tables(0).Rows(I)("Line_Item")
                    '                newRow("Material") = _Material
                    '                newRow("Description") = T01.Tables(0).Rows(I)("Met_Des")
                    '                newRow("Batch Qty (Mtr)") = Microsoft.VisualBasic.Format(T01.Tables(0).Rows(I)("PRD_Qty"), "#.00")
                    '                _QtyMtr = _QtyMtr + T01.Tables(0).Rows(I)("PRD_Qty")
                    '                newRow("Delivary Date") = Month(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Microsoft.VisualBasic.Day(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Year(T01.Tables(0).Rows(I)("Del_Date"))
                    '                _delivaryDT = Month(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Microsoft.VisualBasic.Day(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Year(T01.Tables(0).Rows(I)("Del_Date"))
                    '                newRow("Batch No") = T01.Tables(0).Rows(I)("Prduct_Order")
                    '                newRow("Customer") = T01.Tables(0).Rows(I)("Customer")

                    '                Sql = "select * from ZPP_DEL where Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and Line_Item='" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & "' and Delivery_Date='" & T01.Tables(0).Rows(I)("Del_Date") & "' and Product_Order='" & Trim(T01.Tables(0).Rows(I)("Prduct_Order")) & "'"
                    '                dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    '                If isValidDataset(dsUser) Then
                    '                    newRow("Batch Qty (Kg)") = Microsoft.VisualBasic.Format(dsUser.Tables(0).Rows(0)("Order_Qty_Kg"), "#.00")
                    '                    _QtyKg = _QtyKg + dsUser.Tables(0).Rows(0)("Order_Qty_Kg")
                    '                    newRow("No.Of.Dys In S/L") = dsUser.Tables(0).Rows(0)("No_Day_Same_Opp")
                    '                    newRow("NC Comment") = dsUser.Tables(0).Rows(0)("NC_Comment")
                    '                    newRow("Next Operation") = dsUser.Tables(0).Rows(0)("Status")
                    '                End If

                    '                Sql = "select * from T07OTD_Comment where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07Material='" & _Material & "' and T07BatchNo='" & T01.Tables(0).Rows(I)("Prduct_Order") & "' order by T07Date DESC"
                    '                _OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    '                If isValidDataset(_OTD) Then
                    '                    If IsDate(_OTD.Tables(0).Rows(0)("T07Posible_Date")) Then
                    '                        If Year(_OTD.Tables(0).Rows(0)("T07Posible_Date")) = "1900" Then
                    '                        Else
                    '                            newRow("Possible Devivary Date") = _OTD.Tables(0).Rows(0)("T07Posible_Date")

                    '                        End If
                    '                    End If

                    '                    newRow("Reason") = _OTD.Tables(0).Rows(0)("T07Reason")
                    '                    If _OTD.Tables(0).Rows(0)("T07date") = Today Then
                    '                        newRow(_Todayupdate) = _OTD.Tables(0).Rows(0)("T07Comment")
                    '                    Else
                    '                        newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")
                    '                    End If
                    '                    newRow("LIB dep") = _OTD.Tables(0).Rows(0)("T07LIB")

                    '                    _YesDate = Today.AddDays(-1)
                    '                    Sql = "select * from T07OTD_Comment where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07BatchNo='" & T01.Tables(0).Rows(I)("Prduct_Order") & "' and T07Date>'" & Today & "' order by T07Date DESC"
                    '                    _OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    '                    If isValidDataset(_OTD) Then
                    '                        newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")
                    '                    End If

                    '                Else
                    '                    Sql = "select * from T07OTD_Comment where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07Material='" & _Material & "'  order by T07Date DESC"
                    '                    _OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    '                    If isValidDataset(_OTD) Then
                    '                        '               newRow("Possible Devivary Date") = _OTD.Tables(0).Rows(0)("T07Posible_Date")
                    '                        '              newRow("Reason") = _OTD.Tables(0).Rows(0)("T07Reason")
                    '                        newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")
                    '                        '             newRow("LIB dep") = _OTD.Tables(0).Rows(0)("T07LIB")
                    '                    End If
                    '                End If
                    '                newRow("Week") = "Week " & DatePart(DateInterval.WeekOfYear, _delivaryDT)
                    '                c_dataCustomer1.Rows.Add(newRow)
                    '                ' End If
                    '                _Rowcount = _Rowcount + 1
                    '                ' If Not x = M01.Tables(0).Rows.Count - 1 Then
                    '                ' rsT24RecHeader.MoveNext()
                    '            End If

                    '            ' End If
                    '        End If

                    '    End If
                    'End If
                    '------------------------------------------------------------------------
                    'NEW CODING
                    Dim _Material As String
                    _Material = T01.Tables(0).Rows(I)("Metrrial")
                    '_Material = Microsoft.VisualBasic.Right(_Material, 7)
                    '_Material = Microsoft.VisualBasic.Left(_Material, 2) & "-" & Microsoft.VisualBasic.Right(_Material, 5)
                    newRow("Sales Order") = T01.Tables(0).Rows(I)("Sales_Order")
                    newRow("Line Item") = T01.Tables(0).Rows(I)("Line_Item")
                    newRow("Material") = _Material
                    newRow("Description") = T01.Tables(0).Rows(I)("Met_Des")
                    newRow("Batch Qty (Mtr)") = Microsoft.VisualBasic.Format(T01.Tables(0).Rows(I)("PRD_Qty"), "#.00")
                    _QtyMtr = _QtyMtr + T01.Tables(0).Rows(I)("PRD_Qty")
                    newRow("Delivary Date") = Month(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Microsoft.VisualBasic.Day(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Year(T01.Tables(0).Rows(I)("Del_Date"))
                    _delivaryDT = Month(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Microsoft.VisualBasic.Day(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Year(T01.Tables(0).Rows(I)("Del_Date"))
                    newRow("Batch No") = T01.Tables(0).Rows(I)("Prduct_Order")
                    newRow("Customer") = T01.Tables(0).Rows(I)("Customer")
                    newRow("NC Comment") = T01.Tables(0).Rows(I)("NC_Comment")

                    Sql = "select * from ZPP_DEL where Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and Line_Item='" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & "' and Delivery_Date='" & T01.Tables(0).Rows(I)("Del_Date") & "' and Product_Order='" & Trim(T01.Tables(0).Rows(I)("Prduct_Order")) & "'"
                    dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(dsUser) Then
                        newRow("Batch Qty (Kg)") = Microsoft.VisualBasic.Format(dsUser.Tables(0).Rows(0)("Order_Qty_Kg"), "#.00")
                        _QtyKg = _QtyKg + dsUser.Tables(0).Rows(0)("Order_Qty_Kg")
                        newRow("No.Of.Dys In S/L") = dsUser.Tables(0).Rows(0)("No_Day_Same_Opp")
                        ' newRow("NC Comment") = dsUser.Tables(0).Rows(0)("NC_Comment")
                        newRow("Next Operation") = dsUser.Tables(0).Rows(0)("Status")
                    End If

                    Sql = "select * from T07OTD_Comment1 where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07Material='" & _Material & "' and T07BatchNo='" & T01.Tables(0).Rows(I)("Prduct_Order") & "' order by T07Date DESC"
                    _OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(_OTD) Then
                        If IsDate(_OTD.Tables(0).Rows(0)("T07Posible_Date")) Then
                            If Year(_OTD.Tables(0).Rows(0)("T07Posible_Date")) = "1900" Then
                            Else
                                newRow("Possible Devivary Date") = _OTD.Tables(0).Rows(0)("T07Posible_Date")

                            End If
                        End If

                        newRow("Reason") = _OTD.Tables(0).Rows(0)("T07Reason")
                        If _OTD.Tables(0).Rows(0)("T07date") = Today Then
                            newRow(_Todayupdate) = _OTD.Tables(0).Rows(0)("T07Comment")
                        Else
                            newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")
                        End If
                        newRow("LIB dep") = _OTD.Tables(0).Rows(0)("T07LIB")

                        _YesDate = Today.AddDays(-1)
                        Sql = "select * from T07OTD_Comment1 where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07BatchNo='" & T01.Tables(0).Rows(I)("Prduct_Order") & "' and T07Date<'" & Today & "' order by T07Date DESC"
                        _OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                        If isValidDataset(_OTD) Then
                            newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")
                        End If

                    Else
                        'Sql = "select * from T07OTD_Comment where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07Material='" & _Material & "'  order by T07Date DESC"
                        '_OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                        'If isValidDataset(_OTD) Then
                        '    '               newRow("Possible Devivary Date") = _OTD.Tables(0).Rows(0)("T07Posible_Date")
                        '    '              newRow("Reason") = _OTD.Tables(0).Rows(0)("T07Reason")
                        '    newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")
                        '    '             newRow("LIB dep") = _OTD.Tables(0).Rows(0)("T07LIB")
                        'End If
                    End If
                    newRow("Week") = "Week " & DatePart(DateInterval.WeekOfYear, _delivaryDT)
                    c_dataCustomer1.Rows.Add(newRow)
            I = I + 1
                Next
            ElseIf Trim(cboPO.Text) = "Aw for 2070 location" Then

                Dim _OTD As DataSet
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2070') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')   and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2070') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')   and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2070') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')"

                    End If
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2070') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')    and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2070') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')   and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2070') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') "

                    End If
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2070') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "')   and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2070') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "')   and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2070') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "')"

                    End If
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then

                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2070') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')   and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2070') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')   and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2070') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')"

                    End If

                ElseIf Trim(txtCustomer.Text) <> "" Then

                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2070') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "')   and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2070') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "')   and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2070') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "')"

                    End If

                ElseIf Trim(txtDepartment.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2070') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Department in ('" & _Department & "')    and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2070') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Department in ('" & _Department & "')   and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2070') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Department in ('" & _Department & "') "

                    End If
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2070') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Merchant in ('" & _Merchant & "')   and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2070') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Merchant in ('" & _Merchant & "')   and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2070') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Merchant in ('" & _Merchant & "')"

                    End If
                Else
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2070') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2070') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2070') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') "

                    End If
                End If

                I = 0
                'T01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetMEETING", New SqlParameter("@cQryType", "DTY"), New SqlParameter("@vcWhereClause", vcWharer))
                'T01 = DBEngin.ExecuteDa
                _Rowcount = 0
                _QtyKg = 0
                _QtyMtr = 0

                For Each DTRow3 As DataRow In T01.Tables(0).Rows
                    Dim newRow As DataRow = c_dataCustomer1.NewRow
                    Dim _Tollaranz As Double

                    'For Each DTRow1 As DataRow In M01.Tables(0).Rows


                    Dim _Material As String
                    _Material = T01.Tables(0).Rows(I)("Metrrial")
                    '_Material = Microsoft.VisualBasic.Right(_Material, 7)
                    '_Material = Microsoft.VisualBasic.Left(_Material, 2) & "-" & Microsoft.VisualBasic.Right(_Material, 5)
                    newRow("Sales Order") = T01.Tables(0).Rows(I)("Sales_Order")
                    newRow("Line Item") = T01.Tables(0).Rows(I)("Line_Item")
                    newRow("Material") = _Material
                    newRow("Description") = T01.Tables(0).Rows(I)("Met_Des")
                    newRow("Batch Qty (Mtr)") = Microsoft.VisualBasic.Format(T01.Tables(0).Rows(I)("PRD_Qty"), "#.00")
                    _QtyMtr = _QtyMtr + T01.Tables(0).Rows(I)("PRD_Qty")
                    newRow("Delivary Date") = Month(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Microsoft.VisualBasic.Day(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Year(T01.Tables(0).Rows(I)("Del_Date"))
                    _delivaryDT = Month(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Microsoft.VisualBasic.Day(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Year(T01.Tables(0).Rows(I)("Del_Date"))
                    newRow("Batch No") = T01.Tables(0).Rows(I)("Prduct_Order")
                    newRow("Customer") = T01.Tables(0).Rows(I)("Customer")

                    Sql = "select * from ZPP_DEL where Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and Line_Item='" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & "' and Delivery_Date='" & T01.Tables(0).Rows(I)("Del_Date") & "' and Product_Order='" & Trim(T01.Tables(0).Rows(I)("Prduct_Order")) & "'"
                    dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(dsUser) Then
                        newRow("Batch Qty (Kg)") = Microsoft.VisualBasic.Format(dsUser.Tables(0).Rows(0)("Order_Qty_Kg"), "#.00")
                        _QtyKg = _QtyKg + dsUser.Tables(0).Rows(0)("Order_Qty_Kg")
                        newRow("No.Of.Dys In S/L") = dsUser.Tables(0).Rows(0)("No_Day_Same_Opp")
                        newRow("NC Comment") = dsUser.Tables(0).Rows(0)("NC_Comment")
                        '   newRow("Next Operation") = dsUser.Tables(0).Rows(0)("Status")
                    End If

                    Sql = "select * from T07OTD_Comment1 where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07Material='" & _Material & "' and T07BatchNo='" & T01.Tables(0).Rows(I)("Prduct_Order") & "' order by T07Date DESC"
                    _OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(_OTD) Then
                        If IsDate(_OTD.Tables(0).Rows(0)("T07Posible_Date")) Then
                            If Year(_OTD.Tables(0).Rows(0)("T07Posible_Date")) = "1900" Then
                            Else
                                newRow("Possible Devivary Date") = _OTD.Tables(0).Rows(0)("T07Posible_Date")

                            End If
                        End If

                        newRow("Reason") = _OTD.Tables(0).Rows(0)("T07Reason")
                        If _OTD.Tables(0).Rows(0)("T07date") = Today Then
                            newRow(_Todayupdate) = _OTD.Tables(0).Rows(0)("T07Comment")
                        Else
                            newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")
                        End If
                        newRow("LIB dep") = _OTD.Tables(0).Rows(0)("T07LIB")

                        _YesDate = Today.AddDays(-1)
                        Sql = "select * from T07OTD_Comment1 where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07BatchNo='" & T01.Tables(0).Rows(I)("Prduct_Order") & "' and T07Date<'" & Today & "' order by T07Date DESC"
                        _OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                        If isValidDataset(_OTD) Then
                            newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")
                        End If

                    Else
                        'Sql = "select * from T07OTD_Comment where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07Material='" & _Material & "'  order by T07Date DESC"
                        '_OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                        'If isValidDataset(_OTD) Then
                        '    'newRow("Possible Devivary Date") = _OTD.Tables(0).Rows(0)("T07Posible_Date")
                        '    'newRow("Reason") = _OTD.Tables(0).Rows(0)("T07Reason")
                        '    newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")
                        '    'newRow("LIB dep") = _OTD.Tables(0).Rows(0)("T07LIB")
                        'End If
                    End If

                    newRow("Week") = "Week " & DatePart(DateInterval.WeekOfYear, _delivaryDT)
                    c_dataCustomer1.Rows.Add(newRow)
                    ' End If
                    _Rowcount = _Rowcount + 1
                    ' If Not x = M01.Tables(0).Rows.Count - 1 Then
                    ' rsT24RecHeader.MoveNext()

                    I = I + 1
                Next

            ElseIf Trim(cboPO.Text) = "Aw Quality" Then

                Dim _OTD As DataSet
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Exam' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')   and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Exam' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')   and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Exam' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')"

                    End If
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Exam' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')    and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Exam' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')    and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Exam' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') "

                    End If
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Exam' and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "')   and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Exam' and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "')   and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Exam' and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "')"

                    End If
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then

                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Exam' and  r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')   and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Exam' and  r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')   and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Exam' and and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')"

                    End If

                ElseIf Trim(txtCustomer.Text) <> "" Then

                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Exam' and s.Customer in ('" & _Customer & "')    and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Exam' and s.Customer in ('" & _Customer & "')    and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Exam' and s.Customer in ('" & _Customer & "') "

                    End If

                ElseIf Trim(txtDepartment.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Exam' and  r.Department in ('" & _Department & "')   and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Exam' and  r.Department in ('" & _Department & "')   and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Exam' and  r.Department in ('" & _Department & "') "

                    End If
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Exam' and  r.Merchant in ('" & _Merchant & "')   and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Exam' and  r.Merchant in ('" & _Merchant & "')   and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Exam' and  r.Merchant in ('" & _Merchant & "')"

                    End If
                Else
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Exam'  and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Exam'    and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Exam' "

                    End If
                End If

                I = 0
                'T01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetMEETING", New SqlParameter("@cQryType", "GRG"), New SqlParameter("@vcWhereClause", vcWharer))
                'T01 = DBEngin.ExecuteDa
              
                _Rowcount = 0
                _QtyKg = 0
                _QtyMtr = 0
                For Each DTRow3 As DataRow In T01.Tables(0).Rows
                    Dim newRow As DataRow = c_dataCustomer1.NewRow
                    Dim _Tollaranz As Double

                    'For Each DTRow1 As DataRow In M01.Tables(0).Rows


                    Dim _Material As String
                    _Material = T01.Tables(0).Rows(I)("Metrrial")
                    '_Material = Microsoft.VisualBasic.Right(_Material, 7)
                    '_Material = Microsoft.VisualBasic.Left(_Material, 2) & "-" & Microsoft.VisualBasic.Right(_Material, 5)
                    newRow("Sales Order") = T01.Tables(0).Rows(I)("Sales_Order")
                    newRow("Line Item") = T01.Tables(0).Rows(I)("Line_Item")
                    newRow("Material") = _Material
                    newRow("Description") = T01.Tables(0).Rows(I)("Met_Des")
                    newRow("Batch Qty (Mtr)") = Microsoft.VisualBasic.Format(T01.Tables(0).Rows(I)("PRD_Qty"), "#.00")
                    _QtyMtr = _QtyMtr + T01.Tables(0).Rows(I)("PRD_Qty")
                    newRow("Delivary Date") = Month(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Microsoft.VisualBasic.Day(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Year(T01.Tables(0).Rows(I)("Del_Date"))
                    _delivaryDT = Month(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Microsoft.VisualBasic.Day(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Year(T01.Tables(0).Rows(I)("Del_Date"))

                    newRow("Batch No") = T01.Tables(0).Rows(I)("Prduct_Order")
                    newRow("Customer") = T01.Tables(0).Rows(I)("Customer")


                    Dim _NO_SameLocation As Integer
                    _NO_SameLocation = 0

                    Sql = "select * from ZPP_DEL where Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and Line_Item='" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & "' and Delivery_Date='" & T01.Tables(0).Rows(I)("Del_Date") & "' and Product_Order='" & Trim(T01.Tables(0).Rows(I)("Prduct_Order")) & "'"
                    dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(dsUser) Then
                        newRow("Batch Qty (Kg)") = Microsoft.VisualBasic.Format(dsUser.Tables(0).Rows(0)("Order_Qty_Kg"), "#.00")
                        _QtyKg = _QtyKg + dsUser.Tables(0).Rows(0)("Order_Qty_Kg")
                        newRow("No.Of.Dys In S/L") = dsUser.Tables(0).Rows(0)("No_Day_Same_Opp")
                        _NO_SameLocation = dsUser.Tables(0).Rows(0)("No_Day_Same_Opp")
                        newRow("Next Operation") = dsUser.Tables(0).Rows(0)("Status")
                        If Trim(dsUser.Tables(0).Rows(0)("Status")) = "EXAM & LAB" Then
                        Else
                            newRow("NC Comment") = dsUser.Tables(0).Rows(0)("NC_Comment")
                        End If
                    End If

                    Sql = "select * from T07OTD_Comment1 where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07Material='" & _Material & "' and T07BatchNo='" & T01.Tables(0).Rows(I)("Prduct_Order") & "' order by T07Date DESC"
                    _OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(_OTD) Then
                        If IsDate(_OTD.Tables(0).Rows(0)("T07Posible_Date")) Then
                            If Year(_OTD.Tables(0).Rows(0)("T07Posible_Date")) = "1900" Then
                            Else
                                newRow("Possible Devivary Date") = _OTD.Tables(0).Rows(0)("T07Posible_Date")

                            End If
                        End If

                        newRow("Reason") = _OTD.Tables(0).Rows(0)("T07Reason")
                        If _OTD.Tables(0).Rows(0)("T07date") = Today Then
                            newRow(_Todayupdate) = _OTD.Tables(0).Rows(0)("T07Comment")
                        Else
                            newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")
                        End If
                        newRow("LIB dep") = _OTD.Tables(0).Rows(0)("T07LIB")

                        _YesDate = Today.AddDays(-1)
                        Sql = "select * from T07OTD_Comment1 where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07BatchNo='" & T01.Tables(0).Rows(I)("Prduct_Order") & "' and T07Date<'" & Today & "' order by T07Date DESC"
                        _OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                        If isValidDataset(_OTD) Then
                            newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")
                        End If

                    Else
                        'Sql = "select * from T07OTD_Comment where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07Material='" & _Material & "'  order by T07Date DESC"
                        '_OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                        'If isValidDataset(_OTD) Then
                        '    'newRow("Possible Devivary Date") = _OTD.Tables(0).Rows(0)("T07Posible_Date")
                        '    'newRow("Reason") = _OTD.Tables(0).Rows(0)("T07Reason")
                        '    newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")
                        '    'newRow("LIB dep") = _OTD.Tables(0).Rows(0)("T07LIB")
                        'End If
                    End If

                    newRow("Week") = "Week " & DatePart(DateInterval.WeekOfYear, _delivaryDT)
                    c_dataCustomer1.Rows.Add(newRow)
                    ' End If
                    Dim _Rcount As Integer
                    _Rcount = UltraGrid4.Rows.Count
                    If _NO_SameLocation >= 3 Then
                        UltraGrid4.Rows(_Rcount - 1).Cells(0).Appearance.BackColor = Color.Yellow
                        UltraGrid4.Rows(_Rcount - 1).Cells(1).Appearance.BackColor = Color.Yellow
                        UltraGrid4.Rows(_Rcount - 1).Cells(2).Appearance.BackColor = Color.Yellow
                        UltraGrid4.Rows(_Rcount - 1).Cells(3).Appearance.BackColor = Color.Yellow
                        UltraGrid4.Rows(_Rcount - 1).Cells(4).Appearance.BackColor = Color.Yellow
                        UltraGrid4.Rows(_Rcount - 1).Cells(5).Appearance.BackColor = Color.Yellow
                        UltraGrid4.Rows(_Rcount - 1).Cells(6).Appearance.BackColor = Color.Yellow
                        UltraGrid4.Rows(_Rcount - 1).Cells(7).Appearance.BackColor = Color.Yellow
                        UltraGrid4.Rows(_Rcount - 1).Cells(8).Appearance.BackColor = Color.Yellow
                        UltraGrid4.Rows(_Rcount - 1).Cells(9).Appearance.BackColor = Color.Yellow
                        UltraGrid4.Rows(_Rcount - 1).Cells(10).Appearance.BackColor = Color.Yellow
                        UltraGrid4.Rows(_Rcount - 1).Cells(11).Appearance.BackColor = Color.Yellow
                        UltraGrid4.Rows(_Rcount - 1).Cells(12).Appearance.BackColor = Color.Yellow
                        UltraGrid4.Rows(_Rcount - 1).Cells(13).Appearance.BackColor = Color.Yellow
                        UltraGrid4.Rows(_Rcount - 1).Cells(14).Appearance.BackColor = Color.Yellow
                        UltraGrid4.Rows(_Rcount - 1).Cells(15).Appearance.BackColor = Color.Yellow
                        UltraGrid4.Rows(_Rcount - 1).Cells(16).Appearance.BackColor = Color.Yellow

                        UltraGrid4.Rows(_Rcount - 1).Cells(17).Appearance.BackColor = Color.Yellow
                    End If
                    ' If Not x = M01.Tables(0).Rows.Count - 1 Then
                    ' rsT24RecHeader.MoveNext()
                    _Rowcount = _Rowcount + 1
                    I = I + 1
                Next
            ElseIf Trim(cboPO.Text) = "Aw Print" Then

                Dim _OTD As DataSet
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Aw Print' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')   and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Aw Print' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')   and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Aw Print' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')"

                    End If
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Aw Print' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')    and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Aw Print' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')    and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Aw Print' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') "

                    End If
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Aw Print' and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "')   and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Aw Print' and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "')   and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Aw Print' and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "')"

                    End If
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then

                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Aw Print' and  r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')   and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Aw Print' and  r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')   and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Aw Print' and and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')"

                    End If

                ElseIf Trim(txtCustomer.Text) <> "" Then

                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Aw Print' and s.Customer in ('" & _Customer & "')    and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Aw Print' and s.Customer in ('" & _Customer & "')    and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Aw Print' and s.Customer in ('" & _Customer & "') "

                    End If

                ElseIf Trim(txtDepartment.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Aw Print' and  r.Department in ('" & _Department & "')   and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Aw Print' and  r.Department in ('" & _Department & "')   and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Aw Print' and  r.Department in ('" & _Department & "') "

                    End If
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Aw Print' and  r.Merchant in ('" & _Merchant & "')   and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Aw Print' and  r.Merchant in ('" & _Merchant & "')   and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Aw Print' and  r.Merchant in ('" & _Merchant & "')"

                    End If
                Else
                    If cboQuality.Text <> "" Then
                        If Microsoft.VisualBasic.Left(Trim(cboQuality.Text), 1) = "Y" Then
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Aw Print'  and left(r.Met_Des,8)='" & cboQuality.Text & "'"
                        Else
                            vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Aw Print'    and left(r.Met_Des,5)='" & cboQuality.Text & "'"
                        End If
                    Else
                        vcWharer = " s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Aw Print' "

                    End If
                End If

                I = 0
                'T01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetMEETING", New SqlParameter("@cQryType", "GRG"), New SqlParameter("@vcWhereClause", vcWharer))
                'T01 = DBEngin.ExecuteDa

                _Rowcount = 0
                _QtyKg = 0
                _QtyMtr = 0
                For Each DTRow3 As DataRow In T01.Tables(0).Rows
                    Dim newRow As DataRow = c_dataCustomer1.NewRow
                    Dim _Tollaranz As Double

                    'For Each DTRow1 As DataRow In M01.Tables(0).Rows


                    Dim _Material As String
                    _Material = T01.Tables(0).Rows(I)("Metrrial")
                    '_Material = Microsoft.VisualBasic.Right(_Material, 7)
                    '_Material = Microsoft.VisualBasic.Left(_Material, 2) & "-" & Microsoft.VisualBasic.Right(_Material, 5)
                    newRow("Sales Order") = T01.Tables(0).Rows(I)("Sales_Order")
                    newRow("Line Item") = T01.Tables(0).Rows(I)("Line_Item")
                    newRow("Material") = _Material
                    newRow("Description") = T01.Tables(0).Rows(I)("Met_Des")
                    newRow("Batch Qty (Mtr)") = Microsoft.VisualBasic.Format(T01.Tables(0).Rows(I)("PRD_Qty"), "#.00")
                    _QtyMtr = _QtyMtr + T01.Tables(0).Rows(I)("PRD_Qty")
                    newRow("Delivary Date") = Month(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Microsoft.VisualBasic.Day(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Year(T01.Tables(0).Rows(I)("Del_Date"))
                    _delivaryDT = Month(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Microsoft.VisualBasic.Day(T01.Tables(0).Rows(I)("Del_Date")) & "/" & Year(T01.Tables(0).Rows(I)("Del_Date"))

                    newRow("Batch No") = T01.Tables(0).Rows(I)("Prduct_Order")
                    newRow("Customer") = T01.Tables(0).Rows(I)("Customer")


                    Dim _NO_SameLocation As Integer
                    _NO_SameLocation = 0

                    Sql = "select * from ZPP_DEL where Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and Line_Item='" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & "' and Delivery_Date='" & T01.Tables(0).Rows(I)("Del_Date") & "' and Product_Order='" & Trim(T01.Tables(0).Rows(I)("Prduct_Order")) & "'"
                    dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(dsUser) Then
                        newRow("Batch Qty (Kg)") = Microsoft.VisualBasic.Format(dsUser.Tables(0).Rows(0)("Order_Qty_Kg"), "#.00")
                        _QtyKg = _QtyKg + dsUser.Tables(0).Rows(0)("Order_Qty_Kg")
                        newRow("No.Of.Dys In S/L") = dsUser.Tables(0).Rows(0)("No_Day_Same_Opp")
                        _NO_SameLocation = dsUser.Tables(0).Rows(0)("No_Day_Same_Opp")
                        newRow("Next Operation") = dsUser.Tables(0).Rows(0)("Status")
                        If Trim(dsUser.Tables(0).Rows(0)("Status")) = "EXAM & LAB" Then
                        Else
                            newRow("NC Comment") = dsUser.Tables(0).Rows(0)("NC_Comment")
                        End If
                    End If

                    Sql = "select * from T07OTD_Comment1 where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07Material='" & _Material & "' and T07BatchNo='" & T01.Tables(0).Rows(I)("Prduct_Order") & "' order by T07Date DESC"
                    _OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(_OTD) Then
                        If IsDate(_OTD.Tables(0).Rows(0)("T07Posible_Date")) Then
                            If Year(_OTD.Tables(0).Rows(0)("T07Posible_Date")) = "1900" Then
                            Else
                                newRow("Possible Devivary Date") = _OTD.Tables(0).Rows(0)("T07Posible_Date")

                            End If
                        End If

                        newRow("Reason") = _OTD.Tables(0).Rows(0)("T07Reason")
                        If _OTD.Tables(0).Rows(0)("T07date") = Today Then
                            newRow(_Todayupdate) = _OTD.Tables(0).Rows(0)("T07Comment")
                        Else
                            newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")
                        End If
                        newRow("LIB dep") = _OTD.Tables(0).Rows(0)("T07LIB")

                        _YesDate = Today.AddDays(-1)
                        Sql = "select * from T07OTD_Comment1 where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07BatchNo='" & T01.Tables(0).Rows(I)("Prduct_Order") & "' and T07Date<'" & Today & "' order by T07Date DESC"
                        _OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                        If isValidDataset(_OTD) Then
                            newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")
                        End If

                    Else
                        'Sql = "select * from T07OTD_Comment where T07Sales_Order='" & Trim(T01.Tables(0).Rows(I)("Sales_Order")) & "' and T07Line_Item=" & Trim(T01.Tables(0).Rows(I)("Line_Item")) & " and T07Material='" & _Material & "'  order by T07Date DESC"
                        '_OTD = DBEngin.ExecuteDataset(con, Nothing, Sql)
                        'If isValidDataset(_OTD) Then
                        '    'newRow("Possible Devivary Date") = _OTD.Tables(0).Rows(0)("T07Posible_Date")
                        '    'newRow("Reason") = _OTD.Tables(0).Rows(0)("T07Reason")
                        '    newRow(_Yesterday) = _OTD.Tables(0).Rows(0)("T07Comment")
                        '    'newRow("LIB dep") = _OTD.Tables(0).Rows(0)("T07LIB")
                        'End If
                    End If

                    newRow("Week") = "Week " & DatePart(DateInterval.WeekOfYear, _delivaryDT)
                    c_dataCustomer1.Rows.Add(newRow)
                    ' End If
                    Dim _Rcount As Integer
                    _Rcount = UltraGrid4.Rows.Count
                    If _NO_SameLocation >= 3 Then
                        UltraGrid4.Rows(_Rcount - 1).Cells(0).Appearance.BackColor = Color.Yellow
                        UltraGrid4.Rows(_Rcount - 1).Cells(1).Appearance.BackColor = Color.Yellow
                        UltraGrid4.Rows(_Rcount - 1).Cells(2).Appearance.BackColor = Color.Yellow
                        UltraGrid4.Rows(_Rcount - 1).Cells(3).Appearance.BackColor = Color.Yellow
                        UltraGrid4.Rows(_Rcount - 1).Cells(4).Appearance.BackColor = Color.Yellow
                        UltraGrid4.Rows(_Rcount - 1).Cells(5).Appearance.BackColor = Color.Yellow
                        UltraGrid4.Rows(_Rcount - 1).Cells(6).Appearance.BackColor = Color.Yellow
                        UltraGrid4.Rows(_Rcount - 1).Cells(7).Appearance.BackColor = Color.Yellow
                        UltraGrid4.Rows(_Rcount - 1).Cells(8).Appearance.BackColor = Color.Yellow
                        UltraGrid4.Rows(_Rcount - 1).Cells(9).Appearance.BackColor = Color.Yellow
                        UltraGrid4.Rows(_Rcount - 1).Cells(10).Appearance.BackColor = Color.Yellow
                        UltraGrid4.Rows(_Rcount - 1).Cells(11).Appearance.BackColor = Color.Yellow
                        UltraGrid4.Rows(_Rcount - 1).Cells(12).Appearance.BackColor = Color.Yellow
                        UltraGrid4.Rows(_Rcount - 1).Cells(13).Appearance.BackColor = Color.Yellow
                        UltraGrid4.Rows(_Rcount - 1).Cells(14).Appearance.BackColor = Color.Yellow
                        UltraGrid4.Rows(_Rcount - 1).Cells(15).Appearance.BackColor = Color.Yellow
                        UltraGrid4.Rows(_Rcount - 1).Cells(16).Appearance.BackColor = Color.Yellow

                        UltraGrid4.Rows(_Rcount - 1).Cells(17).Appearance.BackColor = Color.Yellow
                    End If
                    ' If Not x = M01.Tables(0).Rows.Count - 1 Then
                    ' rsT24RecHeader.MoveNext()
                    _Rowcount = _Rowcount + 1
                    I = I + 1
                Next
            End If
            Dim newRow1 As DataRow = c_dataCustomer1.NewRow
            newRow1("Batch Qty (Kg)") = Microsoft.VisualBasic.Format(_QtyKg, "#.00")
            newRow1("Batch Qty (Mtr)") = Microsoft.VisualBasic.Format(_QtyMtr, "#.00")
            c_dataCustomer1.Rows.Add(newRow1)


            'I = 0
            'For Each uRow As UltraGridRow In UltraGrid4.Rows
            '    With UltraGrid4
            '        Sql = "select * from T07OTD_Comment where T07Sales_Order='" & .Rows(I).Cells(0).Value & "' and T07Line_Item='" & .Rows(I).Cells(1).Value & "' and T07BatchNo='" & .Rows(I).Cells(5).Value & "' order by T07Date decs"
            '        T01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            '        If isValidDataset(T01) Then
            '            If T01.Tables(0).Rows(0)("T07Date") = Today Then
            '                .Rows(I).Cells(12).Value = T01.Tables(0).Rows(0)("T07Posible_Date")
            '                .Rows(I).Cells(13).Value = T01.Tables(0).Rows(0)("T07Posible_Date")
            '                .Rows(I).Cells(14).Value = T01.Tables(0).Rows(0)("T07Posible_Date")
            '                .Rows(I).Cells(15).Value = T01.Tables(0).Rows(0)("T07Posible_Date")
            '            Else
            '                .Rows(I).Cells(11).Value = T01.Tables(0).Rows(0)("T07Posible_Date")
            '            End If
            '        Else
            '            Sql = "select * from T07OTD_Comment where T07Sales_Order='" & .Rows(I).Cells(0).Value & "' and T07Line_Item='" & .Rows(I).Cells(1).Value & "'"
            '            T01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            '            If isValidDataset(T01) Then
            '                .Rows(I).Cells(11).Value = T01.Tables(0).Rows(0)("T07Posible_Date")
            '            End If
            '        End If
            '    End With
            '    I = I + 1
            'Next
            _Rowcount = UltraGrid4.Rows.Count
            If _Rowcount > 1 Then
                UltraGrid4.Rows(_Rowcount - 1).Cells(6).Appearance.BackColor = Color.DeepSkyBlue
                UltraGrid4.Rows(_Rowcount - 1).Cells(7).Appearance.BackColor = Color.DeepSkyBlue
            End If
            con.close()

            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            ' worksheet1.Cells(4, 5) = _Fail_Batch
            'worksheet1.Cells(4, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
            ' MsgBox("Report Genarated successfully", MsgBoxStyle.Information, "Technova ....")
            ' MsgBox(_Fail_Batch)
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
                con.close()
                '  MsgBox(I)
            End If
        End Try
    End Function

    Private Sub UltraGrid4_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles UltraGrid4.InitializeLayout
        e.Layout.Bands(0).Columns("Reason").ValueList = Me.UltraDropDown3
        e.Layout.Bands(0).Columns("LIB dep").ValueList = Me.UltraDropDown1
    End Sub

    Private Sub cboPO_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles cboPO.InitializeLayout

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

            dt.Columns.Add("##", GetType(String))
            Sql = "select * from M11Delay_Reason"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            I = 0
            'cmdSave.Enabled = True

            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                dt.Rows.Add(New Object() {M01.Tables(0).Rows(I)("M11Reason")})
                ' dt.Rows.Add(New Object() {"NOT APPROVED"})
                ' dt.Rows.Add(New Object() {"SESANAL"})
                dt.AcceptChanges()
                I = I + 1
            Next
            Me.UltraDropDown3.SetDataBinding(dt, Nothing)
            '  Me.UltraDropDown1.ValueMember = "ID"
            Me.UltraDropDown3.DisplayMember = "##"
            Me.UltraDropDown3.Rows.Band.Columns(0).Width = 275

            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Sub

    Private Sub BindUltraDropDown1()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim I As Integer
        Dim dt As DataTable = New DataTable()
        ' dt.Columns.Add("ID", GetType(Integer))
        Try

            dt.Columns.Add("##", GetType(String))
            Sql = "select * from M12LIB_Department"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            I = 0
            'cmdSave.Enabled = True

            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                dt.Rows.Add(New Object() {M01.Tables(0).Rows(I)("M12Department")})
                ' dt.Rows.Add(New Object() {"NOT APPROVED"})
                ' dt.Rows.Add(New Object() {"SESANAL"})
                dt.AcceptChanges()
                I = I + 1
            Next
            Me.UltraDropDown1.SetDataBinding(dt, Nothing)
            '  Me.UltraDropDown1.ValueMember = "ID"
            Me.UltraDropDown1.DisplayMember = "##"
            Me.UltraDropDown1.Rows.Band.Columns(0).Width = 125

            con.close()
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        Call Load_Main_Gride()
        Call Load_Data_To_GrideNew()
    End Sub

    Private Sub UltraButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton1.Click
        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Dim i As Integer

        Try
            i = 0
            For Each uRow As UltraGridRow In UltraGrid4.Rows
                'QUARANTINE REASON FOR REPORT TABLE
                With UltraGrid4
                    nvcFieldList1 = "select * from T07OTD_Comment1 where T07Sales_Order='" & .Rows(i).Cells(0).Value & "' and T07Line_Item='" & .Rows(i).Cells(1).Value & "' and T07Date='" & Today & "' and T07BatchNo='" & .Rows(i).Cells(5).Value & "' and T07Status='N'"
                    dsUser = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                    If isValidDataset(dsUser) Then
                        If Trim(.Rows(i).Cells(16).Text) <> "" And Trim(.Rows(i).Cells(13).Text) <> "" Or Trim(.Rows(i).Cells(12).Text) <> "" Or Trim(.Rows(i).Cells(14).Text) <> "" Or Trim(.Rows(i).Cells(15).Text) <> "" Then

                            nvcFieldList1 = "update T07OTD_Comment1 set T07Posible_Date='" & Trim(.Rows(i).Cells(14).Text) & "',T07Comment='" & Trim(.Rows(i).Cells(12).Text) & "',T07Reason='" & .Rows(i).Cells(15).Value & "',T07LIB='" & .Rows(i).Cells(16).Value & "',T07Complete_Date='" & .Rows(i).Cells(13).Value & "',T07Status='N' where T07Sales_Order='" & .Rows(i).Cells(0).Value & "' and T07Line_Item='" & .Rows(i).Cells(1).Value & "' and T07Date='" & Today & "' and T07BatchNo='" & .Rows(i).Cells(5).Value & "'"
                            ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                        End If
                    Else
                        If Trim(.Rows(i).Cells(16).Text) <> "" And Trim(.Rows(i).Cells(13).Text) <> "" Or Trim(.Rows(i).Cells(12).Text) <> "" Or Trim(.Rows(i).Cells(14).Text) <> "" Or Trim(.Rows(i).Cells(15).Text) <> "" Then
                            nvcFieldList1 = "Insert Into T07OTD_Comment1(T07Sales_Order,T07Line_Item,T07Material,T07BatchNo,T07Date,T07Posible_Date,T07Comment,T07Reason,T07LIB,T07Complete_Date,T07Status)" & _
                                                   " values('" & .Rows(i).Cells(0).Value & "','" & .Rows(i).Cells(1).Value & "','" & .Rows(i).Cells(2).Value & "','" & .Rows(i).Cells(5).Value & "','" & Today & "','" & .Rows(i).Cells(14).Value & "','" & .Rows(i).Cells(12).Value & "','" & .Rows(i).Cells(15).Value & "','" & .Rows(i).Cells(16).Value & "','" & .Rows(i).Cells(13).Value & "','N')"
                            ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                        End If

                    End If
                End With
                i = i + 1
            Next
            MsgBox("Records successfully updated", MsgBoxStyle.Information, "Information ..........")

            transaction.Commit()
            Load_Main_Gride()
            connection.Close()
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
            common.ClearAll(OPR0, OPR2)
            OPR0.Enabled = True
            OPR2.Enabled = True
            connection.Close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
                connection.Close()
            End If
        End Try

    End Sub

    Private Sub UltraButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton2.Click
        Dim i As Integer
        Dim sr As System.IO.StreamReader
        Dim strFileName As String
        Dim BatchNo As String
        Dim _Coment As String
        Dim _PossibleDate As Date
        Dim _Reason As String
        Dim _LabDip As String
        Dim x1 As Integer

        Try

            x1 = 0
            strFileName = ConfigurationManager.AppSettings("MtnUpload") + "\OTDMtn_Update.txt"
            For Each line As String In System.IO.File.ReadAllLines(strFileName)
                Dim fields() As String = line.Split(vbTab)
                BatchNo = Trim(fields(0))
                _Coment = Trim(fields(1))
                If IsDate(Trim(fields(2))) Then
                    _PossibleDate = Trim(fields(2))
                End If
                _Reason = Trim(fields(3))
                _LabDip = Trim(fields(4))

                i = 0

                For Each uRow As UltraGridRow In UltraGrid4.Rows
                    'QUARANTINE REASON FOR REPORT TABLE
                    'If UltraGrid4.Rows(i).Cells(0).Text <> "" Then

                    'Else
                    '    MsgBox("")
                    'End If
                    With UltraGrid4
                        
                        If .Rows(i).Cells(5).Text = BatchNo Then
                            UltraGrid4.Rows(i).Cells(12).Value = _Coment
                            If IsDate(_PossibleDate) Then
                                UltraGrid4.Rows(i).Cells(14).Value = _PossibleDate
                            End If
                            UltraGrid4.Rows(i).Cells(15).Value = _Reason
                            UltraGrid4.Rows(i).Cells(16).Value = _LabDip
                        End If
                        ' End If
                    End With

                    i = i + 1
                Next

                x1 = x1 + 1
            Next

            ' Con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                'DBEngin.CloseConnection(con)
                'con.ConnectionString = ""
                'Con.close()
                ' MsgBox(x1)
            End If
        End Try
    End Sub

    Private Sub UltraButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton6.Click
        cboS_Order.Text = ""
        txtBatch.Text = ""
        OPR10.Visible = False
    End Sub

    Private Sub cboPO_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboPO.KeyUp
        If e.KeyCode = Keys.F1 Then
            OPR10.Visible = True

        End If
    End Sub

    Private Sub cboQuality_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboQuality.KeyUp
        If e.KeyCode = Keys.F1 Then
            OPR10.Visible = True

        End If
    End Sub

    Private Sub UltraButton7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton7.Click
        Dim i As Integer
        Try
            If cboS_Order.Text <> "" And txtBatch.Text <> "" Then

            ElseIf cboS_Order.Text <> "" Then
                i = 0
                For Each uRow As UltraGridRow In UltraGrid4.Rows
                    If Trim(cboS_Order.Text) = UltraGrid4.Rows(i).Cells(0).Value Then
                        'UltraGrid4.Focus()
                        UltraGrid4.DisplayLayout.Rows(i).Selected = True

                        'UltraGrid4.Rows(i).Cells(0).SelLength = Len(UltraGrid4.Rows(i).Cells(0).Text)
                        'UltraGrid4.Rows(i).Cells(0).Text=
                        ' UltraGrid4.Rows(i).Cells(0).SelLength = Len(UltraGrid4.Rows(i).Cells(0).Value)
                        Exit For
                    End If
                    i = i + 1
                Next

            ElseIf txtBatch.Text <> "" Then
                i = 0
                For Each uRow As UltraGridRow In UltraGrid3.Rows

                Next
            End If

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                ' MsgBox(i)
                'DBEngin.CloseConnection(con)
                'con.ConnectionString = ""
                'con.Close()
            End If
        End Try
    End Sub


    Private Sub UltraLabel7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraLabel7.Click

    End Sub
End Class