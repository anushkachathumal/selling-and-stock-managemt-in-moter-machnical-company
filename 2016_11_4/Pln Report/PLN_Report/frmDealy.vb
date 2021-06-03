
Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports System.IO.StreamReader
Imports Microsoft.Office.Interop.Excel
Imports System.Globalization.Calendar

Public Class frmDealy
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim c_dataCustomer1 As System.Data.DataTable
    Dim _Customer As String
    Dim _Department As String
    Dim _Merchant As String

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

    Private Sub frmDealy_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtFromDate.Text = Today
        txtTodate.Text = Today

        ' Call Load_Customer()
        'Call Load_Merchant()

        'Call Load_Department()
        ' Call Load_Status()
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

    Function Upload_zpp_delFile()
        Dim sr As System.IO.StreamReader
        Dim strFileName As String

        Dim _sales_Order As String
        Dim _LineItem As String

        Dim _Material As String
        Dim _Material_Dis As String
     
        Dim _Del_Date As Date
        Dim _Order_QtyMtr As Double
        Dim _Order_QtyKg As Double

      
        Dim _DealyDate As Integer
        Dim _LastConf_Date As String
        Dim _NoofDays As Integer
        Dim _WIPLoc As String
        Dim _ProductOrder As String
        Dim _Status As String
        Dim _PlnComment As String
        Dim _NCComment As String
        Dim _Dealy_Reason As String
        Dim _OrderType As String
        Dim _Brandix_WIP As String
        Dim _Location As String

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
        Dim characterToRemove As String

        Dim X11 As Integer
        Dim _TollPLS As Integer
        Dim _TollMIN As Integer

        Try
            nvcFieldList1 = "delete from zpp_del"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            strFileName = ConfigurationManager.AppSettings("FilePath") + "\zpp_del.txt"
            For Each line As String In System.IO.File.ReadAllLines(strFileName)
                Dim fields() As String = line.Split(vbTab)

                '  MsgBox(Trim(fields(0)))
                '_Location = Trim(fields(15))
                ' If _Location <> "" Then
                If X11 = 1107 Then
                    '  MsgBox("")
                End If
                If (Trim(fields(1))) <> "" Then
                    _sales_Order = (Trim(fields(1)))
                    _LineItem = CInt(Trim(fields(2)))

                    _Material = Trim(fields(5))
                    _Material_Dis = Trim(fields(6))
                    characterToRemove = "'"

                    'MsgBox(Trim(fields(9)))
                    _Material_Dis = (Replace(_Material_Dis, characterToRemove, ""))
                    Dim B As String
                    B = Microsoft.VisualBasic.Left(Trim(fields(8)), 6)
                    _Del_Date = (Microsoft.VisualBasic.Right(B, 2) & "/" & Microsoft.VisualBasic.Right(Trim(fields(8)), 2) & "/" & Microsoft.VisualBasic.Left(B, 4))

                    _DealyDate = CInt(Trim(fields(7)))

                    ' _Merchnat = Trim(fields(8))

                    ' _Del_Date = Trim(fields(8))
                    _LastConf_Date = Trim(fields(9))
                    '_Del_Date = Trim(fields(9))
                    _NoofDays = Trim(fields(10))
                    _WIPLoc = Trim(fields(11))
                    _ProductOrder = Trim(fields(12))
                    _Order_QtyMtr = Trim(fields(13))
                    _Order_QtyKg = Trim(fields(14))
                    _Status = Trim(fields(15))
                    _PlnComment = Trim(fields(16))
                    _NCComment = Trim(fields(17))
                    _Dealy_Reason = Trim(fields(18))
                    _OrderType = Trim(fields(19))
                    _Brandix_WIP = Trim(fields(20))
                    _Location = Trim(fields(21))

                    characterToRemove = "'"

                    'MsgBox(Trim(fields(9)))
                    _PlnComment = (Replace(_PlnComment, characterToRemove, ""))
                    _Dealy_Reason = (Replace(_Dealy_Reason, characterToRemove, ""))
                    nvcFieldList1 = "select * from ZPP_DEL where Sales_Order='" & Trim(_sales_Order) & "' and Line_Item='" & Trim(_LineItem) & "' and Delivery_Date='" & _Del_Date & "' and Product_Order='" & _ProductOrder & "'"
                    M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                    If isValidDataset(M01) Then
                    Else
                        nvcFieldList1 = "Insert Into ZPP_DEL(Sales_Order,Line_Item,Material,Material_Dis,No_Dealy_Day,Delivery_Date,Lst_Confirn_Date,No_Day_Same_Opp,WIP_Loc,Product_Order,Order_Qty_mtr,Order_Qty_Kg,Status,Pln_Comment,NC_Comment,Delay_reason,Order_Type,Brandix_OTD_Week,Location)" & _
                                                            " values('" & Trim(_sales_Order) & "', '" & Trim(_LineItem) & "','" & Trim(_Material) & "','" & Trim(_Material_Dis) & "','" & _DealyDate & "','" & Trim(_Del_Date) & "','" & Trim(_LastConf_Date) & "','" & Trim(_NoofDays) & "','" & Trim(_WIPLoc) & "','" & _ProductOrder & "','" & _Order_QtyMtr & "','" & _Order_QtyKg & "','" & _Status & "','" & _PlnComment & "','" & _NCComment & "','" & _Dealy_Reason & "','" & _OrderType & "','" & _Brandix_WIP & "','" & _Location & "')"
                        ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                    End If
                    ' pbCount.Value = pbCount.Value + 1


                    lblPro.Text = Trim(fields(1)) & "-" & Trim(fields(2))

                    _sales_Order = ""
                    _LineItem = ""
                    _Material = ""
                    _Material_Dis = ""
                    _Order_QtyMtr = 0
                    _Order_QtyKg = 0
                    _LastConf_Date = ""
                    _NoofDays = 0
                    _WIPLoc = ""
                    _ProductOrder = ""
                    _Status = ""
                    _PlnComment = ""
                    _NCComment = ""
                    _Dealy_Reason = ""
                    _OrderType = ""
                    _Brandix_WIP = ""
                    _Location = ""
                End If
                X11 = X11 + 1
                'cmdEdit.Enabled = True
            Next
            'MsgBox("Record Update successfully", MsgBoxStyle.Information, "Information ....")
            transaction.Commit()
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""



        Catch malFormLineEx As Microsoft.VisualBasic.FileIO.MalformedLineException
            MessageBox.Show("Line " & malFormLineEx.Message & "is not valid and will be skipped.", "Malformed Line Exception")
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
        Catch ex As Exception
            MessageBox.Show(ex.Message & " exception has occurred.", "Exception")
            MsgBox(X11)
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
        Finally
            ' If successful or if an exception is thrown,
            ' close the TextFieldParser.
            ' theTextFieldParser.Close()
        End Try
    End Function

    Function Upload_File1()
        Dim sr As System.IO.StreamReader
        Dim strFileName As String
        Dim _PO_No As String
        Dim _sales_Order As String
        Dim _LineItem As String
        Dim _Shadule As String
        Dim _Material As String
        Dim _Material_Dis As String
        Dim _Customer As String
        Dim _CustomerName As String
        Dim _Department As String
        Dim _Merchnat As String
        Dim _Del_Date As Date
        Dim _Order_Qty As Double
        Dim _Del_Qty As Double
        Dim _Confirm_Qty As Double
        Dim _FGStock As Double
        Dim _Balance As Double
        Dim _Location As String
        Dim _PRD_Qty As String
        Dim _Grg_Qty As Double
        Dim _NCComment As String
        Dim _Awaiting As String
        Dim _depComm As String
        Dim _Comm2 As String
        Dim _OTDStatus As String
        Dim _PRD_OrderQty As Double


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
        Dim characterToRemove As String

        Dim X11 As Integer
        Dim _TollPLS As Integer
        Dim _TollMIN As Integer

        Try


            strFileName = ConfigurationManager.AppSettings("FilePath") + "\delsum.txt"
            For Each line As String In System.IO.File.ReadAllLines(strFileName)
                Dim fields() As String = line.Split(vbTab)
                If X11 = 620 Then
                    ' MsgBox("")
                End If

                '  MsgBox(Trim(fields(0)))
                '_Location = Trim(fields(15))
                ' If _Location <> "" Then
                _PO_No = Trim(fields(3))
                _sales_Order = CInt(Trim(fields(0)))
                _LineItem = CInt(Trim(fields(5)))

                _Material = Trim(fields(7))
                _Material_Dis = Trim(fields(8))
                characterToRemove = "'"

                'MsgBox(Trim(fields(9)))
                _Material_Dis = (Replace(_Material_Dis, characterToRemove, ""))
                _Customer = CInt(Trim(fields(1)))
                _CustomerName = Trim(fields(2))

                _PO_No = (Replace(_PO_No, characterToRemove, ""))

                '   _CustomerName = Microsoft.VisualBasic.Left(_Customer, 2)
                characterToRemove = "'"

                'MsgBox(Trim(fields(9)))
                _CustomerName = (Replace(_CustomerName, characterToRemove, ""))
                _Department = Trim(fields(6))
                If Microsoft.VisualBasic.Left(_Department, 3) = "M&S" Then
                    _Department = Microsoft.VisualBasic.Left(_Department, 3)
                End If
                ' _Merchnat = Trim(fields(8))
                Dim B As String
                B = Microsoft.VisualBasic.Left(Trim(fields(4)), 6)
                _Del_Date = (Microsoft.VisualBasic.Right(B, 2) & "/" & Microsoft.VisualBasic.Right(Trim(fields(4)), 2) & "/" & Microsoft.VisualBasic.Left(B, 4))
                '_Del_Date = Trim(fields(9))
                _Order_Qty = Trim(fields(9))
                _Del_Qty = Trim(fields(12))
                _Confirm_Qty = Trim(fields(10))
                _FGStock = Trim(fields(13))
                _TollPLS = Trim(fields(18))
                _TollMIN = Trim(fields(17))
                _depComm = Trim(fields(14))
                characterToRemove = "'"

                'MsgBox(Trim(fields(9)))
                _depComm = (Replace(_depComm, characterToRemove, ""))


                characterToRemove = ";"

                'MsgBox(Trim(fields(9)))
                _PO_No = (Replace(_PO_No, characterToRemove, ""))
                Dim _Week As Integer

                _Week = DatePart(DateInterval.WeekOfYear, _Del_Date)

                nvcFieldList1 = "select * from M01Sales_Order_SAP where M01Sales_Order='" & Trim(_sales_Order) & "' and M01Line_Item='" & Trim(_LineItem) & "' and M01Material_No='" & _Material & "' "
                M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                If isValidDataset(M01) Then
                Else
                    nvcFieldList1 = "Insert Into M01Sales_Order_SAP(M01Sales_Order,M01PO,M01Customer_Code,M01Cuatomer_Name,M01SO_Date,M01Line_Item,M01Department,M01Material_No,M01Quality,M01SO_Qty,M01Con_Qty,M01Delivary_Qty,M01Cus_Tol_Min,M01Cus_Tol_Pls,M01Tobe_Deliverd,M01Reason_Rejection,M01Status)" & _
                                                        " values('" & Trim(_sales_Order) & "', '" & Trim(_PO_No) & "','" & Trim(_Customer) & "','" & Trim(_CustomerName) & "','" & _Del_Date & "','" & Trim(_LineItem) & "','" & Trim(_Department) & "','" & Trim(_Material) & "','" & Trim(_Material_Dis) & "','" & _Order_Qty & "','" & _Confirm_Qty & "','" & _Del_Qty & "','" & _TollMIN & "','" & _TollPLS & "','" & _FGStock & "','" & _depComm & "','A')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                End If
                ' pbCount.Value = pbCount.Value + 1


                lblPro.Text = Trim(fields(0)) & "-" & Trim(fields(1))
                _PO_No = ""
                _sales_Order = ""
                _LineItem = ""
                '_LineItem = ""
                _Awaiting = ""
                _Balance = 0
                _TollPLS = 0
                _TollMIN = 0
                _Grg_Qty = 0
                _PRD_OrderQty = 0
                _PRD_Qty = ""
                _NCComment = ""
                _Del_Qty = 0
                _Comm2 = ""
                _Customer = ""
                _FGStock = 0
                _depComm = ""
                _Material = ""
                _Material_Dis = ""
                _Merchnat = ""
                _Department = ""
                _Shadule = ""
                _Confirm_Qty = 0
                ' End If
                X11 = X11 + 1
                'cmdEdit.Enabled = True
            Next
            'MsgBox("Record Update successfully", MsgBoxStyle.Information, "Information ....")
            transaction.Commit()
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""



        Catch malFormLineEx As Microsoft.VisualBasic.FileIO.MalformedLineException
            MessageBox.Show("Line " & malFormLineEx.Message & "is not valid and will be skipped.", "Malformed Line Exception")
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
        Catch ex As Exception
            MessageBox.Show(ex.Message & " exception has occurred.", "Exception")
            MsgBox(X11)
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
        Finally
            ' If successful or if an exception is thrown,
            ' close the TextFieldParser.
            ' theTextFieldParser.Close()
        End Try
    End Function

    Function Upload_File()
        Dim sr As System.IO.StreamReader
        Dim strFileName As String
        Dim _PO_No As String
        Dim _sales_Order As String
        Dim _LineItem As String
        Dim _Shadule As String
        Dim _Material As String
        Dim _Material_Dis As String
        Dim _Customer As String
        Dim _CustomerName As String
        Dim _Department As String
        Dim _Merchnat As String
        Dim _Del_Date As Date
        Dim _Order_Qty As Double
        Dim _Del_Qty As Double
        Dim _Confirm_Qty As Double
        Dim _FGStock As Double
        Dim _Balance As Double
        Dim _Location As String
        Dim _PRD_Qty As String
        Dim _Grg_Qty As Double
        Dim _NCComment As String
        Dim _Awaiting As String
        Dim _depComm As String
        Dim _Comm2 As String
        Dim _OTDStatus As String
        Dim _PRD_OrderQty As Double
        Dim _Conqty As Double
        Dim _30Class As String

        Dim QualityNo As String

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
        Dim characterToRemove As String

        Dim X11 As Integer
        Dim _TollPLS As Integer
        Dim _TollMIN As Integer
        Dim _Merchant1 As String
        Dim _Location1 As String
        Dim _Confact As Double

        Try


            strFileName = ConfigurationManager.AppSettings("FilePath") + "\delsum.txt"
            For Each line As String In System.IO.File.ReadAllLines(strFileName)
                Dim fields() As String = line.Split(vbTab)
                If X11 = 10 Then
                    '   MsgBox("")
                End If

                '  MsgBox(Trim(fields(0)))
                '_Location = Trim(fields(15))
                ' If _Location <> "" Then
                _PO_No = Trim(fields(3))
                _sales_Order = CInt(Trim(fields(0)))
                _LineItem = CInt(Trim(fields(5)))

                _Material = Trim(fields(7))
                _Material_Dis = Trim(fields(8))
                characterToRemove = "'"

                'MsgBox(Trim(fields(9)))
                _Material_Dis = (Replace(_Material_Dis, characterToRemove, ""))

                characterToRemove = """"
                _30Class = _Material

                'MsgBox(Trim(fields(9)))
                _Material_Dis = (Replace(_Material_Dis, characterToRemove, ""))

                characterToRemove = "-"
                _30Class = (Replace(_30Class, characterToRemove, ""))

                _Customer = CInt(Trim(fields(1)))
                _CustomerName = Trim(fields(2))

                If Microsoft.VisualBasic.Right(_Material_Dis, 3) = "OCI" Then
                    _Location1 = "OCI"
                ElseIf Microsoft.VisualBasic.Right(_Material_Dis, 3) = "PTL" Then
                    _Location1 = "PTL"
                Else

                    _Location1 = "IN HOUSE"
                End If
                _PO_No = (Replace(_PO_No, characterToRemove, ""))

                '   _CustomerName = Microsoft.VisualBasic.Left(_Customer, 2)
                characterToRemove = "'"

                'MsgBox(Trim(fields(9)))
                _CustomerName = (Replace(_CustomerName, characterToRemove, ""))
                _Department = Trim(fields(6))
                If Microsoft.VisualBasic.Left(_Department, 3) = "M&S" Then
                    _Department = Microsoft.VisualBasic.Left(_Department, 3)
                End If
                ' _Merchnat = Trim(fields(8))

                Dim TestString As String = _Material_Dis
                Dim TestArray() As String = Split(TestString)

                ' TestArray holds {"apple", "", "", "", "pear", "banana", "", ""} 
                Dim LastNonEmpty As Integer = -1
                For z As Integer = 0 To TestArray.Length - 1
                    If TestArray(z) <> "" Then
                        LastNonEmpty += 1
                        TestArray(LastNonEmpty) = TestArray(z)
                        QualityNo = TestArray(LastNonEmpty)
                        Exit For
                    End If
                Next

                Dim B As String
                B = Microsoft.VisualBasic.Left(Trim(fields(4)), 6)
                _Del_Date = (Microsoft.VisualBasic.Right(B, 2) & "/" & Microsoft.VisualBasic.Right(Trim(fields(4)), 2) & "/" & Microsoft.VisualBasic.Left(B, 4))
                '_Del_Date = Trim(fields(9))
                _Order_Qty = Trim(fields(9))
                _Del_Qty = Trim(fields(12))
                _Confirm_Qty = Trim(fields(10))
                _FGStock = Trim(fields(13))
                _TollPLS = Trim(fields(18))
                _TollMIN = Trim(fields(17))
                _Merchant1 = Trim(fields(19))
                If Trim(fields(20)) <> "" Then
                    _Confact = Trim(fields(20))
                Else
                    _Confact = 0
                End If
                _depComm = Trim(fields(14))
                characterToRemove = "'"

                'MsgBox(Trim(fields(9)))
                _depComm = (Replace(_depComm, characterToRemove, ""))


                characterToRemove = ";"

                'MsgBox(Trim(fields(9)))
                _PO_No = (Replace(_PO_No, characterToRemove, ""))

                characterToRemove = "'"

                'MsgBox(Trim(fields(9)))
                _PO_No = (Replace(_PO_No, characterToRemove, ""))
                Dim _Week As Integer

                _Week = DatePart(DateInterval.WeekOfYear, _Del_Date)

                '_Confact = 0
                'nvcFieldList1 = "select * from M22Tec_Spec where M22Quality='" & QualityNo & "'"
                'dsUser = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                'If isValidDataset(dsUser) Then
                '    _Confact = dsUser.Tables(0).Rows(0)("M22Con_Fact")
                'End If


                nvcFieldList1 = "select * from M01Sales_Order_SAP where M01Sales_Order='" & Trim(_sales_Order) & "' and M01Line_Item='" & Trim(_LineItem) & "' and M01Material_No='" & _Material & "' "
                M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                If isValidDataset(M01) Then
                    nvcFieldList1 = "update M01Sales_Order_SAP set M01SO_Qty='" & _Order_Qty & "',M01Con_Qty='" & _Confirm_Qty & "',M01Delivary_Qty='" & _Del_Qty & "',M01SO_Date='" & _Del_Date & "' where M01Sales_Order='" & Trim(_sales_Order) & "' and M01Line_Item='" & Trim(_LineItem) & "' and M01Material_No='" & _Material & "' "
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                Else
                    nvcFieldList1 = "Insert Into M01Sales_Order_SAP(M01Sales_Order,M01PO,M01Customer_Code,M01Cuatomer_Name,M01SO_Date,M01Line_Item,M01Department,M01Material_No,M01Quality,M01SO_Qty,M01Con_Qty,M01Delivary_Qty,M01Cus_Tol_Min,M01Cus_Tol_Pls,M01Tobe_Deliverd,M01Reason_Rejection,M01Status,M01Merchant,M01Quality_No,M01Location,M01Con_Fact,M0130Class)" & _
                                                        " values('" & Trim(_sales_Order) & "', '" & Trim(_PO_No) & "','" & Trim(_Customer) & "','" & Trim(_CustomerName) & "','" & _Del_Date & "','" & Trim(_LineItem) & "','" & Trim(_Department) & "','" & Trim(_Material) & "','" & Trim(_Material_Dis) & "','" & _Order_Qty & "','" & _Confirm_Qty & "','" & _Del_Qty & "','" & _TollMIN & "','" & _TollPLS & "','" & _FGStock & "','" & _depComm & "','A','" & _Merchant1 & "','" & QualityNo & "','" & _Location1 & "','" & _Confact & "','" & _30Class & "')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                End If
                ' pbCount.Value = pbCount.Value + 1


                '  lblPro.Text = Trim(fields(0)) & "-" & Trim(fields(1))
                _PO_No = ""
                _sales_Order = ""
                _LineItem = ""
                '_LineItem = ""
                _Awaiting = ""
                _Balance = 0
                _TollPLS = 0
                _TollMIN = 0
                _Grg_Qty = 0
                _PRD_OrderQty = 0
                _PRD_Qty = ""
                _NCComment = ""
                _Del_Qty = 0
                _Comm2 = ""
                _Customer = ""
                _FGStock = 0
                _depComm = ""
                _Material = ""
                _Material_Dis = ""
                _Merchnat = ""
                _Department = ""
                _Shadule = ""
                _Confirm_Qty = 0
                _Location1 = ""
                _Confact = 0
                ' End If
                X11 = X11 + 1
                'cmdEdit.Enabled = True
            Next
            'MsgBox("Record Update successfully", MsgBoxStyle.Information, "Information ....")
            transaction.Commit()
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""

            connection.Close()
            pbCount.Value = 16
            lblPro.Text = "Delsum.txt"
            lblPro.Refresh()
            pbCount.Refresh()
        Catch malFormLineEx As Microsoft.VisualBasic.FileIO.MalformedLineException
            MessageBox.Show("Line " & malFormLineEx.Message & "is not valid and will be skipped.", "Malformed Line Exception")
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
        Catch ex As Exception
            MessageBox.Show(ex.Message & " exception has occurred.", "Exception")
            MsgBox(X11)
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
            connection.Close()
        Finally
            ' If successful or if an exception is thrown,
            ' close the TextFieldParser.
            ' theTextFieldParser.Close()
        End Try
    End Function

    Function Edit_OTD_Status()
        Dim sr As System.IO.StreamReader
        Dim strFileName As String
        Dim _PO_No As String
        Dim _sales_Order As String
        Dim _LineItem As String
        Dim _Shadule As String
        Dim _Material As String
        Dim _Material_Dis As String
        Dim _Customer As String
        Dim _Department As String
        Dim _Merchnat As String
        Dim _Del_Date As String
        Dim _Order_Qty As Double
        Dim _Del_Qty As Double
        Dim _Delay_Qty As Double
        Dim _FGStock As Double
        Dim _Balance As Double
        Dim _Location As String
        Dim _PRD_Qty As String
        Dim _Grg_Qty As Double
        Dim _NCComment As String
        Dim _Awaiting As String
        Dim _depComm As String
        Dim _Comm2 As String
        Dim _OTDStatus As String
        Dim _PRD_OrderQty As Double



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
        Dim characterToRemove As String

        Dim X11 As Integer
        Dim _CusCode As String

        Try
            nvcFieldList1 = "delete from OTD_SMS"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            strFileName = ConfigurationManager.AppSettings("FilePath") + "\otdStatus.txt"
            For Each line As String In System.IO.File.ReadAllLines(strFileName)
                Dim fields() As String = line.Split(vbTab)
                If X11 = 2 Then
                    ' MsgBox("")
                End If
                ' _Location = Trim(fields(15))
                'If _Location <> "" Then

                _sales_Order = Trim(fields(1))
                _LineItem = Trim(fields(3))
                _CusCode = Trim(fields(0))
                _Customer = Trim(fields(16))
                characterToRemove = "'"

                'MsgBox(Trim(fields(9)))
                _Customer = (Replace(_Customer, characterToRemove, " "))

                Dim B As String
                B = Microsoft.VisualBasic.Left(Trim(fields(8)), 6)
                _Del_Date = (Microsoft.VisualBasic.Right(B, 2) & "/" & Microsoft.VisualBasic.Right(Trim(fields(8)), 2) & "/" & Microsoft.VisualBasic.Left(B, 4))

                ' _Del_Date = Trim(fields(8))


                If Trim(fields(11)) = "1" Then
                    _OTDStatus = "True"
                Else
                    _OTDStatus = "Fales"
                End If
                '_OTDStatus = "Fales"

                characterToRemove = "."

                'MsgBox(Trim(fields(9)))
                '_Del_Date = (Replace(_Del_Date, characterToRemove, "/"))
                Dim A1 As String

                A1 = (Microsoft.VisualBasic.Left(_Del_Date, 5))
                ' _Del_Date = Microsoft.VisualBasic.Right(A1, 2) & "/" & Microsoft.VisualBasic.Left(A1, 2) & "/" & Microsoft.VisualBasic.Right(_Del_Date, 4)
                'Dim oDate As DateTime = Convert.ToDateTime(_Del_Date)
                'MsgBox(oDate.Day & " " & oDate.Month & "  " & oDate.Year)
                _Department = Trim(fields(18))
                If Microsoft.VisualBasic.Left(_Department, 3) = "M&S" Then
                    _Department = Microsoft.VisualBasic.Left(_Department, 3)
                Else

                    _Department = Trim(fields(18))
                End If

                _Merchnat = Trim(fields(17))


                nvcFieldList1 = "select * from OTD_Records where Sales_Order='" & Trim(_sales_Order) & "' and Line_Item='" & Trim(_LineItem) & "' and Del_Date='" & _Del_Date & "' "
                M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                If isValidDataset(M01) Then
                    nvcFieldList1 = "update OTD_Records set Cus_Code='" & _CusCode & "',Customer='" & _Customer & "',Status='" & _OTDStatus & "' where Sales_Order='" & Trim(_sales_Order) & "' and Line_Item='" & _LineItem & "' and Del_Date='" & _Del_Date & "'"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                End If
                ' pbCount.Value = pbCount.Value + 1

                'INSERT OTD_SMS

                nvcFieldList1 = "SELECT * FROM OTD_SMS WHERE Sales_Order='" & Trim(_sales_Order) & "' AND Line_Item='" & _LineItem & "' AND Del_Date='" & _Del_Date & "'"
                M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                If isValidDataset(M01) Then
                Else
                    nvcFieldList1 = "Insert Into OTD_SMS(Sales_Order,Line_Item,Cus_Code,Customer,Del_Date,Status,Department,Merchant)" & _
                                                        " values('" & Trim(_sales_Order) & "', '" & Trim(_LineItem) & "','" & Trim(_CusCode) & "','" & Trim(_Customer) & "','" & _Del_Date & "','" & Trim(_OTDStatus) & "','" & _Department & "','" & _Merchnat & "')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                End If
                _CusCode = ""
                _Customer = ""
                _sales_Order = ""
                _LineItem = ""
                _OTDStatus = ""
                X11 = X11 + 1
                'cmdEdit.Enabled = True
            Next
            ' MsgBox("Record Update successfully", MsgBoxStyle.Information, "Information ....")

            nvcFieldList1 = "select * from OTD_SMS S inner join OTD_Records r on r.Sales_Order=s.Sales_Order and r.Line_Item=s.Line_Item and r.Del_Date=s.Del_Date where s.Status='Fales' and r.Delay_Qty=r.FG_Stock"
            M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            I = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows

                nvcFieldList1 = "update OTD_Records set Status='True' where Sales_Order='" & Trim(M01.Tables(0).Rows(I)("Sales_Order")) & "' and Line_Item='" & Trim(M01.Tables(0).Rows(I)("Line_Item")) & "' and Del_Date='" & M01.Tables(0).Rows(I)("Del_Date") & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                I = I + 1
            Next
            transaction.Commit()
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
        Catch malFormLineEx As Microsoft.VisualBasic.FileIO.MalformedLineException
            MessageBox.Show("Line " & malFormLineEx.Message & "is not valid and will be skipped.", "Malformed Line Exception")
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
        Catch ex As Exception
            MessageBox.Show(ex.Message & " exception has occurred.", "Exception")
            MsgBox(X11)
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
        Finally
            ' If successful or if an exception is thrown,
            ' close the TextFieldParser.
            ' theTextFieldParser.Close()
        End Try
    End Function

    Private Sub UltraButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton2.Click
        Call Upload_File()
        Call Edit_OTD_Status()
        Call Upload_zpp_delFile()

        MsgBox("Record update successfully", MsgBoxStyle.Information, "Information .....")

    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        Call Craete_File()

    End Sub

    Function Craete_File()
        Dim SQL As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim T01 As DataSet
        Dim T03 As DataSet
        Dim dsUser As DataSet
        Dim tblDye As DataSet

        Dim n_Date As Date
        Dim N_Date1 As Date
        Dim FileName As String
        Dim _FirstChr As Integer
        Dim _Possible_Date As Date
        Dim _Last As Integer

        Dim _Tobe As Integer
        Dim _Aw_Dyeing As Integer
        Dim _Aw_Grage As Integer
        Dim _NCApp As Integer
        Dim _NC_KNT As Integer
        Dim _NC_FINISHING As Integer
        Dim _AWCONTAMI As Integer
        Dim _AWPrint As Integer

        Dim _Awprep As Integer
        Dim _Aw_Finishing As Integer
        Dim _Aw_Exam As Integer
        Dim _2065 As Integer
        Dim _2062 As Integer
        Dim _2070 As Integer

        Dim _Nc1 As Integer

        Dim _Replacements As Integer
        Dim _Shortages As Integer
        Dim _Reprocess_Dyeing As Integer
        Dim _Week As String

        _Tobe = 1
        _Aw_Dyeing = 1
        _Replacements = 1
        _Reprocess_Dyeing = 1
        _Shortages = 1
        Try
            Dim exc As New Application

            Dim workbooks As Workbooks = exc.Workbooks
            Dim workbook As _Workbook = workbooks.Add(XlWBATemplate.xlWBATWorksheet)
            Dim sheets As Sheets = workbook.Worksheets
            Dim worksheet1 As _Worksheet = CType(sheets.Item(1), _Worksheet)

            exc.Visible = True
            Dim i As Integer
            Dim _GrandTotal As Integer
            Dim _STGrand As String
            Dim range1 As Range
            Dim _NETTOTAL As Integer
            Dim T04 As DataSet
            Dim n_per As Double
            Dim Y As Integer
            Dim _cOUNT As Integer
            Dim _Fail_Batch As Integer

            _Fail_Batch = 0

            Dim sheets1 As Sheets = workbook.Worksheets
            Dim worksheet2 As _Worksheet = CType(sheets1.Item(1), _Worksheet)
            worksheet1.Rows(2).Font.size = 11
            worksheet1.Rows(2).Font.Bold = True

            If txtTodate.Text > Today Then

                Dim currentCulture As System.Globalization.CultureInfo
                currentCulture = System.Globalization.CultureInfo.CurrentCulture
                Dim weekNum = currentCulture.Calendar.GetWeekOfYear(txtTodate.Text, System.Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday)
                Dim _S As String
                _S = "Week " & weekNum & " Delivery batches in urgent list as at " & Today
                worksheet1.Cells(2, 1) = _S
                worksheet1.Cells(2, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter

            Else
                worksheet1.Cells(2, 1) = "Delay batches in urgent list as at " & Today
                worksheet1.Cells(2, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter

            End If
            worksheet1.Rows(2).Font.size = 18
            worksheet1.Rows(2).Font.name = "Tahoma"
            worksheet1.Rows(2).rowheight = 24.25
            worksheet1.Rows(6).rowheight = 35.25

            worksheet1.Range("a2:e2").MergeCells = True
            worksheet1.Range("a2:e2").VerticalAlignment = XlVAlign.xlVAlignCenter


            worksheet1.Rows(6).Font.size = 9
            worksheet1.Rows(6).Font.name = "Tahoma"
            worksheet1.Rows(6).Font.Bold = True
            worksheet1.Cells(6, 1) = "Sales Order"
            worksheet1.Cells(6, 2) = "Line Item"
            worksheet1.Cells(6, 3) = "Customer Name"
            worksheet1.Cells(6, 4) = "Material"
            worksheet1.Cells(6, 5) = "Material Description"
            worksheet1.Cells(6, 6) = "No of Delay Days"
            worksheet1.Cells(6, 7) = "Delivery Date"
            worksheet1.Cells(6, 8) = "Last Confirmed Date"
            worksheet1.Cells(6, 9) = "No of days in same opparation"
            worksheet1.Cells(6, 10) = "Poduction Order No"
            worksheet1.Cells(6, 11) = "Order Qty(Kg)"
            worksheet1.Cells(6, 12) = "Order Qty(M)"
            worksheet1.Cells(6, 13) = "Status"
            worksheet1.Cells(6, 14) = "Planning Comments"
            worksheet1.Cells(6, 15) = "NC Comments"
            worksheet1.Cells(6, 16) = "Delay reason"
            worksheet1.Cells(6, 17) = "Order Type"
            worksheet1.Cells(6, 18) = "Brandix OTD Week"
            worksheet1.Cells(6, 19) = "WIP Location"


            worksheet1.Columns("A").ColumnWidth = 15
            worksheet1.Columns("B").ColumnWidth = 10
            worksheet1.Columns("C").ColumnWidth = 29
            worksheet1.Columns("D").ColumnWidth = 12
            worksheet1.Columns("E").ColumnWidth = 35
            worksheet1.Columns("F").ColumnWidth = 10
            worksheet1.Columns("M").ColumnWidth = 20
            worksheet1.Columns("G").ColumnWidth = 14
            worksheet1.Columns("H").ColumnWidth = 14
            worksheet1.Columns("N").ColumnWidth = 18
            worksheet1.Columns("J").ColumnWidth = 20
            worksheet1.Columns("K").ColumnWidth = 13
            worksheet1.Columns("L").ColumnWidth = 13
            worksheet1.Columns("O").ColumnWidth = 25
            worksheet1.Columns("P").ColumnWidth = 24
            worksheet1.Columns("R").ColumnWidth = 17
            worksheet1.Columns("S").ColumnWidth = 17

            Dim _Chart As Integer
            Dim X As Integer
            X = 6
            _Chart = 97
            For i = 1 To 19
                worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Interior.Color = RGB(141, 180, 227)
                worksheet1.Cells(X, i).WrapText = True
                worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).MergeCells = True
                worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                worksheet1.Cells(X, i).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                _Chart = _Chart + 1
            Next
            ' X = X + 1

            '==================================================================
            'TO BE PLANNED
            Dim _FROMDATE As Date
            Dim _TODATE As Date

            _FROMDATE = Month(txtFromDate.Text) & "/" & Microsoft.VisualBasic.Day(txtFromDate.Text) & "/" & Year(txtFromDate.Text)
            _TODATE = Month(txtTodate.Text) & "/" & Microsoft.VisualBasic.Day(txtTodate.Text) & "/" & Year(txtTodate.Text)

            If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,s.Customer " & _
                      "INNER JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                      "WHERE Z.WIP_Loc='To Be planned' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND S.Customer IN ('" & _Customer & "') AND S.Department IN ('" & _Department & "') AND S.Merchant IN ('" & _Merchant & "') "
            ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,s.Customer " & _
                      "INNER JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                      "WHERE Z.WIP_Loc='To Be planned' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND S.Customer IN ('" & _Customer & "') AND S.Department IN ('" & _Department & "')  "
            ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                        "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,s.Customer " & _
                        "INNER JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                        "WHERE Z.WIP_Loc='To Be planned' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND S.Customer IN ('" & _Customer & "') AND S.Merchant IN ('" & _Merchant & "')  "
            ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                        "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,s.Customer " & _
                        "INNER JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                        "WHERE Z.WIP_Loc='To Be planned' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND S.Department IN ('" & _Department & "') AND S.Merchant IN ('" & _Merchant & "')  "
            ElseIf Trim(txtCustomer.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                       "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,s.Customer " & _
                      " FROM ZPP_DEL z INNER JOIN  M01Sales_Order_SAP ON Z.Sales_Order=M01Sales_Order AND M01Line_Item=Z.Line_Item " & _
                     "INNER JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                     "WHERE Z.WIP_Loc='To Be planned' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' and s.Customer in ('" & _Customer & "')"

            ElseIf Trim(txtDepartment.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                         "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,s.Customer " & _
                         "INNER JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                         "WHERE Z.WIP_Loc='To Be planned' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "'  AND S.Department IN ('" & _Department & "')  "
            ElseIf Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,s.Customer " & _
                      "INNER JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                      "WHERE Z.WIP_Loc='To Be planned' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "'  AND S.Merchant IN ('" & _Merchant & "') "

            Else
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,s.Customer " & _
                     " FROM ZPP_DEL z INNER JOIN  M01Sales_Order_SAP ON Z.Sales_Order=M01Sales_Order AND M01Line_Item=Z.Line_Item " & _
                    "INNER JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                    "WHERE Z.WIP_Loc='To Be planned' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "'"
            End If

            X = X + 1
            _FirstChr = X
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(T01) Then
                worksheet1.Cells(X, 1) = "To Be planned"
                worksheet1.Range("A" & X & ":B" & X).MergeCells = True
                worksheet1.Range("A" & X & ":B" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                '  worksheet1.Range("A" & X, "A" & X).Interior.Color = RGB(141, 180, 227)
                worksheet1.Rows(X).Font.size = 10
                worksheet1.Rows(X).Font.name = "Tahoma"
                worksheet1.Rows(X).Font.Bold = True

                _Chart = 97
                For Y = 1 To 19


                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                    _Chart = _Chart + 1
                Next
                X = X + 1
                '  End If
                i = 0
                For Each DTRow3 As DataRow In T01.Tables(0).Rows
                    Dim _Tollaranz As Double
                    worksheet1.Rows(X).Font.size = 8
                    worksheet1.Rows(X).Font.name = "Tahoma"

                    _Tollaranz = 0
                    _Tollaranz = T01.Tables(0).Rows(i)("M01SO_Qty") * (T01.Tables(0).Rows(i)("M01Cus_Tol_Min") / 100)
                    '   MsgBox(CDbl(T01.Tables(0).Rows(i)("Order_Qty_mtr")))
                    If _Tollaranz < CDbl(T01.Tables(0).Rows(i)("Order_Qty_mtr")) Then
                        worksheet1.Cells(X, 1) = T01.Tables(0).Rows(i)("Sales_Order")
                        worksheet1.Cells(X, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        worksheet1.Cells(X, 2) = T01.Tables(0).Rows(i)("Line_Item")
                        worksheet1.Cells(X, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        worksheet1.Cells(X, 3) = T01.Tables(0).Rows(i)("Customer")
                        worksheet1.Cells(X, 4) = T01.Tables(0).Rows(i)("Material")
                        worksheet1.Cells(X, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        worksheet1.Cells(X, 5) = T01.Tables(0).Rows(i)("Material_Dis")
                        worksheet1.Cells(X, 6) = T01.Tables(0).Rows(i)("No_Dealy_Day")
                        worksheet1.Cells(X, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        worksheet1.Cells(X, 7) = T01.Tables(0).Rows(i)("Delivery_Date")
                        worksheet1.Cells(X, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        worksheet1.Cells(X, 8) = T01.Tables(0).Rows(i)("Lst_Confirn_Date")
                        worksheet1.Cells(X, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        worksheet1.Cells(X, 9) = T01.Tables(0).Rows(i)("No_Day_Same_Opp")
                        worksheet1.Cells(X, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        worksheet1.Cells(X, 10) = T01.Tables(0).Rows(i)("Product_Order")
                        worksheet1.Cells(X, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        worksheet1.Cells(X, 11) = T01.Tables(0).Rows(i)("Order_Qty_Kg")
                        ' worksheet1.Cells(X, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        range1 = worksheet1.Cells(X, 11)
                        range1.NumberFormat = "0.00"
                        worksheet1.Cells(X, 12) = T01.Tables(0).Rows(i)("Order_Qty_mtr")
                        ' worksheet1.Cells(X, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        range1 = worksheet1.Cells(X, 11)
                        range1.NumberFormat = "0.00"
                        worksheet1.Cells(X, 13) = T01.Tables(0).Rows(i)("Status")
                        worksheet1.Cells(X, 14) = T01.Tables(0).Rows(i)("Pln_Comment")
                        worksheet1.Cells(X, 15) = T01.Tables(0).Rows(i)("NC_Comment")
                        worksheet1.Cells(X, 16) = T01.Tables(0).Rows(i)("Delay_reason")
                        worksheet1.Cells(X, 17) = T01.Tables(0).Rows(i)("Order_Type")
                        worksheet1.Cells(X, 17).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        worksheet1.Cells(X, 18) = T01.Tables(0).Rows(i)("Brandix_OTD_Week")
                        worksheet1.Cells(X, 19) = T01.Tables(0).Rows(i)("Location")
                        worksheet1.Cells(X, 19).HorizontalAlignment = XlHAlign.xlHAlignCenter

                        _Chart = 97
                        For Y = 1 To 19


                            worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                            worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                            worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                            worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                            _Chart = _Chart + 1
                        Next

                        X = X + 1
                    End If
                    i = i + 1
                Next
                _Tobe = X
                worksheet1.Cells(X, 3) = "Sum of To Be planned"
                worksheet1.Range("C" & X, "C" & X).Interior.Color = RGB(141, 180, 227)
                worksheet1.Rows(X).Font.size = 10
                worksheet1.Rows(X).Font.name = "Tahoma"
                worksheet1.Rows(X).Font.Bold = True

                worksheet1.Range("K" & (X)).Formula = "=SUM(K" & _FirstChr & ":K" & (X - 1) & ")"
                worksheet1.Cells(X, 11).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet1.Cells(X, 11)
                range1.NumberFormat = "0.00"
                worksheet1.Range("K" & X, "K" & X).Interior.Color = RGB(141, 180, 227)

                worksheet1.Range("l" & (X)).Formula = "=SUM(L" & _FirstChr & ":L" & (X - 1) & ")"
                worksheet1.Cells(X, 12).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet1.Cells(X, 12)
                range1.NumberFormat = "0.00"
                worksheet1.Range("L" & X, "L" & X).Interior.Color = RGB(141, 180, 227)

                _Chart = 97
                For Y = 1 To 19


                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                    _Chart = _Chart + 1
                Next

            End If
            '==================================================================================
            'REPROCESS DYE
            If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                      "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                      "WHERE Z.Location='Reprocess – Dyeing' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') AND S.Department IN ('" & _Department & "') AND S.Merchant IN ('" & _Merchant & "') "
            ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                      "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                      "WHERE Z.Location='Reprocess – Dyeing' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') AND S.Department IN ('" & _Department & "')  "
            ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                        "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                        "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                        "WHERE Z.Location='Reprocess – Dyeing' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') AND S.Merchant IN ('" & _Merchant & "')  "
            ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                        "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                        "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                        "WHERE Z.Location='Reprocess – Dyeing' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND S.Department IN ('" & _Department & "') AND S.Merchant IN ('" & _Merchant & "')  "
            ElseIf Trim(txtCustomer.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                                "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                                "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                                "WHERE Z.Location='Reprocess – Dyeing' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') "
            ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                         "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                         "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                         "WHERE Z.Location='Reprocess – Dyeing' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "'  AND S.Department IN ('" & _Department & "')  "
            ElseIf Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                      "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                      "WHERE Z.Location='Reprocess – Dyeing' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "'  AND S.Merchant IN ('" & _Merchant & "') "

            Else
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                     " FROM ZPP_DEL z INNER JOIN  M01Sales_Order_SAP ON Z.Sales_Order=M01Sales_Order AND M01Line_Item=Z.Line_Item " & _
                    "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                    "WHERE Z.Location in ('Reprocess – Dyeing','Aw Dyeing','Aw Finishing','Aw prep for Reprocess') AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' and z.NC_Comment in ('22.RESCHEDULE - RE DYE','23.RESCHEDULE – WASH','24.RESCHEDULE – STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RE SCHEDULE – SAMPLE')"
            End If

            X = X + 1
            _FirstChr = X
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(T01) Then
                worksheet1.Cells(X, 1) = "Reprocess – Dyeing"
                worksheet1.Range("A" & X & ":B" & X).MergeCells = True
                worksheet1.Range("A" & X & ":B" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                '  worksheet1.Range("A" & X, "A" & X).Interior.Color = RGB(141, 180, 227)
                worksheet1.Rows(X).Font.size = 10
                worksheet1.Rows(X).Font.name = "Tahoma"
                worksheet1.Rows(X).Font.Bold = True

                _Chart = 97
                For Y = 1 To 19


                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                    _Chart = _Chart + 1
                Next
                X = X + 1
                _FirstChr = X
                i = 0
                For Each DTRow3 As DataRow In T01.Tables(0).Rows
                    Dim _Tollaranz As Double
                    worksheet1.Rows(X).Font.size = 8
                    worksheet1.Rows(X).Font.name = "Tahoma"

                    _Tollaranz = 0
                    _Tollaranz = T01.Tables(0).Rows(i)("M01SO_Qty") * (T01.Tables(0).Rows(i)("M01Cus_Tol_Min") / 100)
                    '   MsgBox(CDbl(T01.Tables(0).Rows(i)("Order_Qty_mtr")))

                    worksheet1.Cells(X, 1) = T01.Tables(0).Rows(i)("Sales_Order")
                    worksheet1.Cells(X, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 2) = T01.Tables(0).Rows(i)("Line_Item")
                    worksheet1.Cells(X, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 3) = T01.Tables(0).Rows(i)("M01Cuatomer_Name")
                    worksheet1.Cells(X, 4) = T01.Tables(0).Rows(i)("Material")
                    worksheet1.Cells(X, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 5) = T01.Tables(0).Rows(i)("Material_Dis")
                    worksheet1.Cells(X, 6) = T01.Tables(0).Rows(i)("No_Dealy_Day")
                    worksheet1.Cells(X, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 7) = T01.Tables(0).Rows(i)("Delivery_Date")
                    worksheet1.Cells(X, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 8) = T01.Tables(0).Rows(i)("Lst_Confirn_Date")
                    worksheet1.Cells(X, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 9) = T01.Tables(0).Rows(i)("No_Day_Same_Opp")
                    worksheet1.Cells(X, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 10) = T01.Tables(0).Rows(i)("Product_Order")
                    worksheet1.Cells(X, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 11) = T01.Tables(0).Rows(i)("Order_Qty_Kg")
                    ' worksheet1.Cells(X, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(X, 11)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(X, 12) = T01.Tables(0).Rows(i)("Order_Qty_mtr")
                    ' worksheet1.Cells(X, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(X, 11)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(X, 13) = T01.Tables(0).Rows(i)("Status")
                    worksheet1.Cells(X, 14) = T01.Tables(0).Rows(i)("Pln_Comment")
                    worksheet1.Cells(X, 15) = T01.Tables(0).Rows(i)("NC_Comment")
                    worksheet1.Cells(X, 16) = T01.Tables(0).Rows(i)("Delay_reason")
                    worksheet1.Cells(X, 17) = T01.Tables(0).Rows(i)("Order_Type")
                    worksheet1.Cells(X, 17).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 18) = T01.Tables(0).Rows(i)("Brandix_OTD_Week")
                    worksheet1.Cells(X, 19) = T01.Tables(0).Rows(i)("Location")
                    worksheet1.Cells(X, 19).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    _Chart = 97
                    For Y = 1 To 19


                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                        _Chart = _Chart + 1
                    Next

                    X = X + 1

                    i = i + 1
                Next
                _Reprocess_Dyeing = X
                worksheet1.Cells(X, 3) = "Sum of Reprocess – Dyeing"
                worksheet1.Range("C" & X, "C" & X).Interior.Color = RGB(141, 180, 227)
                worksheet1.Rows(X).Font.size = 10
                worksheet1.Rows(X).Font.name = "Tahoma"
                worksheet1.Rows(X).Font.Bold = True

                worksheet1.Range("K" & (X)).Formula = "=SUM(K" & _FirstChr & ":K" & (X - 1) & ")"
                worksheet1.Cells(X, 11).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet1.Cells(X, 11)
                range1.NumberFormat = "0.00"
                worksheet1.Range("K" & X, "K" & X).Interior.Color = RGB(141, 180, 227)

                worksheet1.Range("l" & (X)).Formula = "=SUM(L" & _FirstChr & ":L" & (X - 1) & ")"
                worksheet1.Cells(X, 12).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet1.Cells(X, 12)
                range1.NumberFormat = "0.00"
                worksheet1.Range("L" & X, "L" & X).Interior.Color = RGB(141, 180, 227)

                _Chart = 97
                For Y = 1 To 19

                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                    _Chart = _Chart + 1
                Next
            End If
            '============================================================================================
            'Shortages


            If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                      "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                      "WHERE Z.Location='Shortages' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') AND S.Department IN ('" & _Department & "') AND S.Merchant IN ('" & _Merchant & "') "
            ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                      "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                      "WHERE Z.Location='Shortages' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') AND S.Department IN ('" & _Department & "')  "
            ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                        "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                        "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                        "WHERE Z.Location='Shortages' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') AND S.Merchant IN ('" & _Merchant & "')  "
            ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                        "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                        "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                        "WHERE Z.Location='Shortages' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND S.Department IN ('" & _Department & "') AND S.Merchant IN ('" & _Merchant & "')  "
            ElseIf Trim(txtCustomer.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                                "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                                "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                                "WHERE Z.Location='Shortages' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') "
            ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                         "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                         "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                         "WHERE Z.Location='Shortages' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "'  AND S.Department IN ('" & _Department & "')  "
            ElseIf Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                      "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                      "WHERE Z.Location='Shortages' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "'  AND S.Merchant IN ('" & _Merchant & "') "

            Else
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                     " FROM ZPP_DEL z INNER JOIN  M01Sales_Order_SAP ON Z.Sales_Order=M01Sales_Order AND M01Line_Item=Z.Line_Item " & _
                    "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                    "WHERE Z.Location='Shortages' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "'"
            End If

            X = X + 1
            _FirstChr = X
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(T01) Then
                worksheet1.Cells(X, 1) = "Shortages"
                worksheet1.Range("A" & X & ":B" & X).MergeCells = True
                worksheet1.Range("A" & X & ":B" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                '  worksheet1.Range("A" & X, "A" & X).Interior.Color = RGB(141, 180, 227)
                worksheet1.Rows(X).Font.size = 10
                worksheet1.Rows(X).Font.name = "Tahoma"
                worksheet1.Rows(X).Font.Bold = True

                _Chart = 97
                For Y = 1 To 19


                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                    _Chart = _Chart + 1
                Next
                X = X + 1
                _FirstChr = X
                i = 0
                For Each DTRow3 As DataRow In T01.Tables(0).Rows
                    Dim _Tollaranz As Double
                    worksheet1.Rows(X).Font.size = 8
                    worksheet1.Rows(X).Font.name = "Tahoma"

                    _Tollaranz = 0
                    _Tollaranz = T01.Tables(0).Rows(i)("M01SO_Qty") * (T01.Tables(0).Rows(i)("M01Cus_Tol_Min") / 100)
                    '   MsgBox(CDbl(T01.Tables(0).Rows(i)("Order_Qty_mtr")))

                    worksheet1.Cells(X, 1) = T01.Tables(0).Rows(i)("Sales_Order")
                    worksheet1.Cells(X, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 2) = T01.Tables(0).Rows(i)("Line_Item")
                    worksheet1.Cells(X, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 3) = T01.Tables(0).Rows(i)("M01Cuatomer_Name")
                    worksheet1.Cells(X, 4) = T01.Tables(0).Rows(i)("Material")
                    worksheet1.Cells(X, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 5) = T01.Tables(0).Rows(i)("Material_Dis")
                    worksheet1.Cells(X, 6) = T01.Tables(0).Rows(i)("No_Dealy_Day")
                    worksheet1.Cells(X, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 7) = T01.Tables(0).Rows(i)("Delivery_Date")
                    worksheet1.Cells(X, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 8) = T01.Tables(0).Rows(i)("Lst_Confirn_Date")
                    worksheet1.Cells(X, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 9) = T01.Tables(0).Rows(i)("No_Day_Same_Opp")
                    worksheet1.Cells(X, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 10) = T01.Tables(0).Rows(i)("Product_Order")
                    worksheet1.Cells(X, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 11) = T01.Tables(0).Rows(i)("Order_Qty_Kg")
                    ' worksheet1.Cells(X, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(X, 11)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(X, 12) = T01.Tables(0).Rows(i)("Order_Qty_mtr")
                    ' worksheet1.Cells(X, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(X, 11)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(X, 13) = T01.Tables(0).Rows(i)("Status")
                    worksheet1.Cells(X, 14) = T01.Tables(0).Rows(i)("Pln_Comment")
                    worksheet1.Cells(X, 15) = T01.Tables(0).Rows(i)("NC_Comment")
                    worksheet1.Cells(X, 16) = T01.Tables(0).Rows(i)("Delay_reason")
                    worksheet1.Cells(X, 17) = T01.Tables(0).Rows(i)("Order_Type")
                    worksheet1.Cells(X, 17).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 18) = T01.Tables(0).Rows(i)("Brandix_OTD_Week")
                    worksheet1.Cells(X, 19) = T01.Tables(0).Rows(i)("Location")
                    worksheet1.Cells(X, 19).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    _Chart = 97
                    For Y = 1 To 19


                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                        _Chart = _Chart + 1
                    Next

                    X = X + 1

                    i = i + 1
                Next
                _Shortages = X
                worksheet1.Cells(X, 3) = "Sum of Shortages"
                worksheet1.Range("C" & X, "C" & X).Interior.Color = RGB(141, 180, 227)
                worksheet1.Rows(X).Font.size = 10
                worksheet1.Rows(X).Font.name = "Tahoma"
                worksheet1.Rows(X).Font.Bold = True

                worksheet1.Range("K" & (X)).Formula = "=SUM(K" & _FirstChr & ":K" & (X - 1) & ")"
                worksheet1.Cells(X, 11).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet1.Cells(X, 11)
                range1.NumberFormat = "0.00"
                worksheet1.Range("K" & X, "K" & X).Interior.Color = RGB(141, 180, 227)

                worksheet1.Range("l" & (X)).Formula = "=SUM(L" & _FirstChr & ":L" & (X - 1) & ")"
                worksheet1.Cells(X, 12).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet1.Cells(X, 12)
                range1.NumberFormat = "0.00"
                worksheet1.Range("L" & X, "L" & X).Interior.Color = RGB(141, 180, 227)

                _Chart = 97
                For Y = 1 To 19

                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                    _Chart = _Chart + 1
                Next
            End If
            '===============================================================================================
            'Replacements
            If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,s.Customer,M01Cuatomer_Name " & _
                      "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                      "WHERE Z.Location='Replacements' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') AND S.Department IN ('" & _Department & "') AND S.Merchant IN ('" & _Merchant & "') "
            ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,s.Customer,M01Cuatomer_Name " & _
                      "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                      "WHERE Z.Location='Replacements' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') AND S.Department IN ('" & _Department & "')  "
            ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                        "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,s.Customer,M01Cuatomer_Name " & _
                        "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                        "WHERE Z.Location='Replacements' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') AND S.Merchant IN ('" & _Merchant & "')  "
            ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                        "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,s.Customer,M01Cuatomer_Name " & _
                        "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                        "WHERE Z.Location='Replacements' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND S.Department IN ('" & _Department & "') AND S.Merchant IN ('" & _Merchant & "')  "
            ElseIf Trim(txtCustomer.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                                "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,s.Customer,M01Cuatomer_Name " & _
                                "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                                "WHERE Z.Location='Replacements' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') "
            ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                         "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,s.Customer,M01Cuatomer_Name " & _
                         "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                         "WHERE Z.Location='Replacements' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "'  AND S.Department IN ('" & _Department & "')  "
            ElseIf Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,s.Customer,M01Cuatomer_Name " & _
                      "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                      "WHERE Z.Location='Replacements' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "'  AND S.Merchant IN ('" & _Merchant & "') "

            Else
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,s.Customer,M01Cuatomer_Name " & _
                     " FROM ZPP_DEL z INNER JOIN  M01Sales_Order_SAP ON Z.Sales_Order=M01Sales_Order AND M01Line_Item=Z.Line_Item " & _
                    "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                    "WHERE Z.Location='Replacements' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "'"
            End If

            X = X + 1
            _FirstChr = X
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(T01) Then
                worksheet1.Cells(X, 1) = "Replacements"
                worksheet1.Range("A" & X & ":B" & X).MergeCells = True
                worksheet1.Range("A" & X & ":B" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                '  worksheet1.Range("A" & X, "A" & X).Interior.Color = RGB(141, 180, 227)
                worksheet1.Rows(X).Font.size = 10
                worksheet1.Rows(X).Font.name = "Tahoma"
                worksheet1.Rows(X).Font.Bold = True

                _Chart = 97
                For Y = 1 To 19


                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                    _Chart = _Chart + 1
                Next
                X = X + 1
                _FirstChr = X
                i = 0
                For Each DTRow3 As DataRow In T01.Tables(0).Rows
                    Dim _Tollaranz As Double
                    worksheet1.Rows(X).Font.size = 8
                    worksheet1.Rows(X).Font.name = "Tahoma"

                    _Tollaranz = 0
                    _Tollaranz = T01.Tables(0).Rows(i)("M01SO_Qty") * (T01.Tables(0).Rows(i)("M01Cus_Tol_Min") / 100)
                    '   MsgBox(CDbl(T01.Tables(0).Rows(i)("Order_Qty_mtr")))

                    worksheet1.Cells(X, 1) = T01.Tables(0).Rows(i)("Sales_Order")
                    worksheet1.Cells(X, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 2) = T01.Tables(0).Rows(i)("Line_Item")
                    worksheet1.Cells(X, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 3) = T01.Tables(0).Rows(i)("M01Cuatomer_Name")
                    worksheet1.Cells(X, 4) = T01.Tables(0).Rows(i)("Material")
                    worksheet1.Cells(X, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 5) = T01.Tables(0).Rows(i)("Material_Dis")
                    worksheet1.Cells(X, 6) = T01.Tables(0).Rows(i)("No_Dealy_Day")
                    worksheet1.Cells(X, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 7) = T01.Tables(0).Rows(i)("Delivery_Date")
                    worksheet1.Cells(X, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 8) = T01.Tables(0).Rows(i)("Lst_Confirn_Date")
                    worksheet1.Cells(X, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 9) = T01.Tables(0).Rows(i)("No_Day_Same_Opp")
                    worksheet1.Cells(X, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 10) = T01.Tables(0).Rows(i)("Product_Order")
                    worksheet1.Cells(X, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 11) = T01.Tables(0).Rows(i)("Order_Qty_Kg")
                    ' worksheet1.Cells(X, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(X, 11)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(X, 12) = T01.Tables(0).Rows(i)("Order_Qty_mtr")
                    ' worksheet1.Cells(X, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(X, 11)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(X, 13) = T01.Tables(0).Rows(i)("Status")
                    worksheet1.Cells(X, 14) = T01.Tables(0).Rows(i)("Pln_Comment")
                    worksheet1.Cells(X, 15) = T01.Tables(0).Rows(i)("NC_Comment")
                    worksheet1.Cells(X, 16) = T01.Tables(0).Rows(i)("Delay_reason")
                    worksheet1.Cells(X, 17) = T01.Tables(0).Rows(i)("Order_Type")
                    worksheet1.Cells(X, 17).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 18) = T01.Tables(0).Rows(i)("Brandix_OTD_Week")
                    worksheet1.Cells(X, 19) = T01.Tables(0).Rows(i)("Location")
                    worksheet1.Cells(X, 19).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    _Chart = 97
                    For Y = 1 To 19


                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                        _Chart = _Chart + 1
                    Next

                    X = X + 1

                    i = i + 1
                Next

                _Replacements = X
                worksheet1.Cells(X, 3) = "Sum of Replacements"
                worksheet1.Range("C" & X, "C" & X).Interior.Color = RGB(141, 180, 227)
                worksheet1.Rows(X).Font.size = 10
                worksheet1.Rows(X).Font.name = "Tahoma"
                worksheet1.Rows(X).Font.Bold = True

                worksheet1.Range("K" & (X)).Formula = "=SUM(K" & _FirstChr & ":K" & (X - 1) & ")"
                worksheet1.Cells(X, 11).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet1.Cells(X, 11)
                range1.NumberFormat = "0.00"
                worksheet1.Range("K" & X, "K" & X).Interior.Color = RGB(141, 180, 227)

                worksheet1.Range("l" & (X)).Formula = "=SUM(L" & _FirstChr & ":L" & (X - 1) & ")"
                worksheet1.Cells(X, 12).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet1.Cells(X, 12)
                range1.NumberFormat = "0.00"
                worksheet1.Range("L" & X, "L" & X).Interior.Color = RGB(141, 180, 227)

                _Chart = 97
                For Y = 1 To 19

                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                    _Chart = _Chart + 1
                Next
            End If
            '==================================================================================
            'Aw Dyeing
            'Remove s.Del_Date=z.Delivery_Date request bt Amila on 03/03/2016
            If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,s.Customer,M01Cuatomer_Name " & _
                      "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item  " & _
                      "WHERE Z.Location='Aw Dyeing' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') AND S.Department IN ('" & _Department & "') AND S.Merchant IN ('" & _Merchant & "') "
            ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,s.Customer,M01Cuatomer_Name " & _
                      "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item  " & _
                      "WHERE Z.Location='Aw Dyeing' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') AND S.Department IN ('" & _Department & "')  "
            ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                        "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,s.Customer,M01Cuatomer_Name " & _
                        "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and  " & _
                        "WHERE Z.Location='Aw Dyeing' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') AND S.Merchant IN ('" & _Merchant & "')  "
            ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                        "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,s.Customer,M01Cuatomer_Name " & _
                        "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item  " & _
                        "WHERE Z.Location='Aw Dyeing' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND S.Department IN ('" & _Department & "') AND S.Merchant IN ('" & _Merchant & "')  "
            ElseIf Trim(txtCustomer.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                                "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,s.Customer,M01Cuatomer_Name " & _
                                "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item  " & _
                                "WHERE Z.Location='Aw Dyeing' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') "
            ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                         "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,s.Customer,M01Cuatomer_Name " & _
                         "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item  " & _
                         "WHERE Z.Location='Aw Dyeing' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "'  AND S.Department IN ('" & _Department & "')  "
            ElseIf Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,s.Customer,M01Cuatomer_Name " & _
                      "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item  " & _
                      "WHERE Z.Location='Aw Dyeing' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "'  AND S.Merchant IN ('" & _Merchant & "') "

            Else
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                         "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,s.Customer,M01Cuatomer_Name " & _
                        " FROM ZPP_DEL z INNER JOIN  M01Sales_Order_SAP ON Z.Sales_Order=M01Sales_Order AND M01Line_Item=Z.Line_Item " & _
                       "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item  " & _
                       "WHERE Z.Location='Aw Dyeing' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' and z.NC_Comment not in ('22.RESCHEDULE - RE DYE','23.RESCHEDULE – WASH','24.RESCHEDULE – STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RE SCHEDULE – SAMPLE')"

                'SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                '       "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,s.Customer,M01Cuatomer_Name " & _
                '      " FROM ZPP_DEL z INNER JOIN  M01Sales_Order_SAP ON Z.Sales_Order=M01Sales_Order AND M01Line_Item=Z.Line_Item " & _
                '                  "WHERE Z.Location='Aw Dyeing' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' and z.NC_Comment not in ('22.RESCHEDULE - RE DYE','23.RESCHEDULE – WASH','24.RESCHEDULE – STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RE SCHEDULE – SAMPLE')"

                SQL = "select * from View_Delay_Rpt where Location='Aw Dyeing' and Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' and NC_Comment not in ('22.RESCHEDULE - RE DYE','23.RESCHEDULE – WASH','24.RESCHEDULE – STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RE SCHEDULE – SAMPLE')"
            End If

            X = X + 1
            _FirstChr = X
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(T01) Then
                worksheet1.Cells(X, 1) = "Aw Dyeing"
                worksheet1.Range("A" & X & ":B" & X).MergeCells = True
                worksheet1.Range("A" & X & ":B" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                '  worksheet1.Range("A" & X, "A" & X).Interior.Color = RGB(141, 180, 227)
                worksheet1.Rows(X).Font.size = 10
                worksheet1.Rows(X).Font.name = "Tahoma"
                worksheet1.Rows(X).Font.Bold = True

                _Chart = 97
                For Y = 1 To 19


                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                    _Chart = _Chart + 1
                Next
                X = X + 1
                _FirstChr = X
                i = 0
                For Each DTRow3 As DataRow In T01.Tables(0).Rows
                    Dim _Tollaranz As Double
                    worksheet1.Rows(X).Font.size = 8
                    worksheet1.Rows(X).Font.name = "Tahoma"

                    _Tollaranz = 0
                    _Tollaranz = T01.Tables(0).Rows(i)("M01SO_Qty") * (T01.Tables(0).Rows(i)("M01Cus_Tol_Min") / 100)
                    '   MsgBox(CDbl(T01.Tables(0).Rows(i)("Order_Qty_mtr")))


                    worksheet1.Cells(X, 1) = T01.Tables(0).Rows(i)("Sales_Order")
                    worksheet1.Cells(X, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 2) = T01.Tables(0).Rows(i)("Line_Item")
                    worksheet1.Cells(X, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 3) = T01.Tables(0).Rows(i)("M01Cuatomer_Name")
                    worksheet1.Cells(X, 4) = T01.Tables(0).Rows(i)("Material")
                    worksheet1.Cells(X, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 5) = T01.Tables(0).Rows(i)("Material_Dis")
                    worksheet1.Cells(X, 6) = T01.Tables(0).Rows(i)("No_Dealy_Day")
                    worksheet1.Cells(X, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 7) = T01.Tables(0).Rows(i)("Delivery_Date")
                    worksheet1.Cells(X, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 8) = T01.Tables(0).Rows(i)("Lst_Confirn_Date")
                    worksheet1.Cells(X, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 9) = T01.Tables(0).Rows(i)("No_Day_Same_Opp")
                    worksheet1.Cells(X, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 10) = T01.Tables(0).Rows(i)("Product_Order")
                    worksheet1.Cells(X, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 11) = T01.Tables(0).Rows(i)("Order_Qty_Kg")
                    ' worksheet1.Cells(X, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(X, 11)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(X, 12) = T01.Tables(0).Rows(i)("Order_Qty_mtr")
                    ' worksheet1.Cells(X, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(X, 11)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(X, 13) = T01.Tables(0).Rows(i)("Status")
                    worksheet1.Cells(X, 14) = T01.Tables(0).Rows(i)("Pln_Comment")
                    worksheet1.Cells(X, 15) = T01.Tables(0).Rows(i)("NC_Comment")
                    worksheet1.Cells(X, 16) = T01.Tables(0).Rows(i)("Delay_reason")
                    worksheet1.Cells(X, 17) = T01.Tables(0).Rows(i)("Order_Type")
                    worksheet1.Cells(X, 17).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 18) = T01.Tables(0).Rows(i)("Brandix_OTD_Week")
                    worksheet1.Cells(X, 19) = T01.Tables(0).Rows(i)("Location")
                    worksheet1.Cells(X, 19).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    _Chart = 97
                    For Y = 1 To 19


                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                        _Chart = _Chart + 1
                    Next

                    X = X + 1

                    i = i + 1
                Next

                _Aw_Dyeing = X
                worksheet1.Cells(X, 3) = "Sum of Aw Dyeing"
                worksheet1.Range("C" & X, "C" & X).Interior.Color = RGB(141, 180, 227)
                worksheet1.Rows(X).Font.size = 10
                worksheet1.Rows(X).Font.name = "Tahoma"
                worksheet1.Rows(X).Font.Bold = True

                worksheet1.Range("K" & (X)).Formula = "=SUM(K" & _FirstChr & ":K" & (X - 1) & ")"
                worksheet1.Cells(X, 11).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet1.Cells(X, 11)
                range1.NumberFormat = "0.00"
                worksheet1.Range("K" & X, "K" & X).Interior.Color = RGB(141, 180, 227)

                worksheet1.Range("l" & (X)).Formula = "=SUM(L" & _FirstChr & ":L" & (X - 1) & ")"
                worksheet1.Cells(X, 12).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet1.Cells(X, 12)
                range1.NumberFormat = "0.00"
                worksheet1.Range("L" & X, "L" & X).Interior.Color = RGB(141, 180, 227)

                _Chart = 97
                For Y = 1 To 19

                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                    _Chart = _Chart + 1
                Next
            End If
            ' End IF
            '-------------------------------------------------
            X = X + 2
            Dim _TOTALDYE As Integer
            _TOTALDYE = X
            worksheet1.Cells(X, 3) = "Total Dyeing"
            worksheet1.Range("C" & X, "C" & X).Interior.Color = RGB(141, 180, 227)
            worksheet1.Rows(X).Font.size = 10
            worksheet1.Rows(X).Font.name = "Tahoma"
            worksheet1.Rows(X).Font.Bold = True

            worksheet1.Range("K" & (X)).Formula = "=(K" & _Tobe & "+K" & _Aw_Dyeing & "+K" & _Replacements & "+K" & _Reprocess_Dyeing & "+k" & _Shortages & ")"
            worksheet1.Cells(X, 11).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 11)
            range1.NumberFormat = "0.00"
            worksheet1.Range("K" & X, "K" & X).Interior.Color = RGB(141, 180, 227)

            worksheet1.Range("l" & (X)).Formula = "=(L" & _Tobe & "+L" & _Aw_Dyeing & "+L" & _Replacements & "+L" & _Reprocess_Dyeing & "+L" & _Shortages & ")"
            worksheet1.Cells(X, 12).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 12)
            range1.NumberFormat = "0.00"
            worksheet1.Range("L" & X, "L" & X).Interior.Color = RGB(141, 180, 227)

            _Chart = 97
            For Y = 1 To 19
                worksheet1.Range(ChrW(_Chart) & X - 2, ChrW(_Chart) & X - 2).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet1.Range(ChrW(_Chart) & X - 2, ChrW(_Chart) & X - 2).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                worksheet1.Range(ChrW(_Chart) & X - 2, ChrW(_Chart) & X - 2).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                worksheet1.Range(ChrW(_Chart) & X - 2, ChrW(_Chart) & X - 2).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous


                worksheet1.Range(ChrW(_Chart) & X - 1, ChrW(_Chart) & X - 1).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet1.Range(ChrW(_Chart) & X - 1, ChrW(_Chart) & X - 1).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                worksheet1.Range(ChrW(_Chart) & X - 1, ChrW(_Chart) & X - 1).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                worksheet1.Range(ChrW(_Chart) & X - 1, ChrW(_Chart) & X - 1).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                _Chart = _Chart + 1
            Next
            '----------------------------------------------------------------------------------------------------------------------------------
            'AW GREIGE
            If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                      "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                      "WHERE Z.Location='AW Greige' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') AND S.Department IN ('" & _Department & "') AND S.Merchant IN ('" & _Merchant & "') "
            ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                      "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                      "WHERE Z.Location='AW Greige' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') AND S.Department IN ('" & _Department & "')  "
            ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                        "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                        "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                        "WHERE Z.Location='AW Greige' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') AND S.Merchant IN ('" & _Merchant & "')  "
            ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                        "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                        "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                        "WHERE Z.Location='AW Greige' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND S.Department IN ('" & _Department & "') AND S.Merchant IN ('" & _Merchant & "')  "
            ElseIf Trim(txtCustomer.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                                "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                                "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                                "WHERE Z.Location='AW Greige' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') "
            ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                         "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                         "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                         "WHERE Z.Location='AW Greige' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "'  AND S.Department IN ('" & _Department & "')  "
            ElseIf Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                      "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                      "WHERE Z.Location='AW Greige' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "'  AND S.Merchant IN ('" & _Merchant & "') "

            Else
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                     " FROM ZPP_DEL z INNER JOIN  M01Sales_Order_SAP ON Z.Sales_Order=M01Sales_Order AND M01Line_Item=Z.Line_Item " & _
                    "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                    "WHERE Z.Location='AW Greige' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "'"

                SQL = "select * from View_Delay_Rpt where Location='AW Greige' and Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "'"
            End If


            X = X + 1
            _FirstChr = X
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(T01) Then
                worksheet1.Cells(X, 1) = "AW Greige"
                worksheet1.Range("A" & X & ":B" & X).MergeCells = True
                worksheet1.Range("A" & X & ":B" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                '  worksheet1.Range("A" & X, "A" & X).Interior.Color = RGB(141, 180, 227)
                worksheet1.Rows(X).Font.size = 10
                worksheet1.Rows(X).Font.name = "Tahoma"
                worksheet1.Rows(X).Font.Bold = True

                _Chart = 97
                For Y = 1 To 19

                    worksheet1.Range(ChrW(_Chart) & X - 1, ChrW(_Chart) & X - 1).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X - 1, ChrW(_Chart) & X - 1).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X - 1, ChrW(_Chart) & X - 1).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X - 1, ChrW(_Chart) & X - 1).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                    _Chart = _Chart + 1
                Next
                X = X + 1
                _FirstChr = X
                i = 0
                For Each DTRow3 As DataRow In T01.Tables(0).Rows
                    Dim _Tollaranz As Double
                    worksheet1.Rows(X).Font.size = 8
                    worksheet1.Rows(X).Font.name = "Tahoma"

                    _Tollaranz = 0
                    _Tollaranz = T01.Tables(0).Rows(i)("M01SO_Qty") * (T01.Tables(0).Rows(i)("M01Cus_Tol_Min") / 100)
                    '   MsgBox(CDbl(T01.Tables(0).Rows(i)("Order_Qty_mtr")))

                    worksheet1.Cells(X, 1) = T01.Tables(0).Rows(i)("Sales_Order")
                    worksheet1.Cells(X, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 2) = T01.Tables(0).Rows(i)("Line_Item")
                    worksheet1.Cells(X, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 3) = T01.Tables(0).Rows(i)("M01Cuatomer_Name")
                    worksheet1.Cells(X, 4) = T01.Tables(0).Rows(i)("Material")
                    worksheet1.Cells(X, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 5) = T01.Tables(0).Rows(i)("Material_Dis")
                    worksheet1.Cells(X, 6) = T01.Tables(0).Rows(i)("No_Dealy_Day")
                    worksheet1.Cells(X, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 7) = T01.Tables(0).Rows(i)("Delivery_Date")
                    worksheet1.Cells(X, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 8) = T01.Tables(0).Rows(i)("Lst_Confirn_Date")
                    worksheet1.Cells(X, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 9) = T01.Tables(0).Rows(i)("No_Day_Same_Opp")
                    worksheet1.Cells(X, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 10) = T01.Tables(0).Rows(i)("Product_Order")
                    worksheet1.Cells(X, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 11) = T01.Tables(0).Rows(i)("Order_Qty_Kg")
                    ' worksheet1.Cells(X, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(X, 11)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(X, 12) = T01.Tables(0).Rows(i)("Order_Qty_mtr")
                    ' worksheet1.Cells(X, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(X, 11)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(X, 13) = T01.Tables(0).Rows(i)("Status")
                    worksheet1.Cells(X, 14) = T01.Tables(0).Rows(i)("Pln_Comment")
                    worksheet1.Cells(X, 15) = T01.Tables(0).Rows(i)("NC_Comment")
                    worksheet1.Cells(X, 16) = T01.Tables(0).Rows(i)("Delay_reason")
                    worksheet1.Cells(X, 17) = T01.Tables(0).Rows(i)("Order_Type")
                    worksheet1.Cells(X, 17).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 18) = T01.Tables(0).Rows(i)("Brandix_OTD_Week")
                    worksheet1.Cells(X, 19) = T01.Tables(0).Rows(i)("Location")
                    worksheet1.Cells(X, 19).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    _Chart = 97
                    For Y = 1 To 19


                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                        _Chart = _Chart + 1
                    Next

                    X = X + 1

                    i = i + 1
                Next

                _Aw_Grage = X
                worksheet1.Cells(X, 3) = "Sum of AW Greige"
                worksheet1.Range("C" & X, "C" & X).Interior.Color = RGB(141, 180, 227)
                worksheet1.Rows(X).Font.size = 10
                worksheet1.Rows(X).Font.name = "Tahoma"
                worksheet1.Rows(X).Font.Bold = True

                worksheet1.Range("K" & (X)).Formula = "=SUM(K" & _FirstChr & ":K" & (X - 1) & ")"
                worksheet1.Cells(X, 11).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet1.Cells(X, 11)
                range1.NumberFormat = "0.00"
                worksheet1.Range("K" & X, "K" & X).Interior.Color = RGB(141, 180, 227)

                worksheet1.Range("l" & (X)).Formula = "=SUM(L" & _FirstChr & ":L" & (X - 1) & ")"
                worksheet1.Cells(X, 12).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet1.Cells(X, 12)
                range1.NumberFormat = "0.00"
                worksheet1.Range("L" & X, "L" & X).Interior.Color = RGB(141, 180, 227)

                _Chart = 97
                For Y = 1 To 19

                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                    _Chart = _Chart + 1
                Next
            End If
            '-------------------------------------------------------------------------------------------------------------------
            ' N/C Held & waiting for approvel

            If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                      "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                      "WHERE Z.Location='N/C Held & waiting for Approva' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') AND S.Department IN ('" & _Department & "') AND S.Merchant IN ('" & _Merchant & "') "
            ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                      "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                      "WHERE Z.Location='N/C Held & waiting for Approva' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') AND S.Department IN ('" & _Department & "')  "
            ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                        "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                        "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                        "WHERE Z.Location='N/C Held & waiting for Approva' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') AND S.Merchant IN ('" & _Merchant & "')  "
            ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                        "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                        "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                        "WHERE Z.Location='N/C Held & waiting for Approva' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND S.Department IN ('" & _Department & "') AND S.Merchant IN ('" & _Merchant & "')  "
            ElseIf Trim(txtCustomer.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                                "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                                "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                                "WHERE Z.Location='N/C Held & waiting for Approva' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') "
            ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                         "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                         "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                         "WHERE Z.Location='N/C Held & waiting for Approva' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "'  AND S.Department IN ('" & _Department & "')  "
            ElseIf Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                      "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                      "WHERE Z.Location='N/C Held & waiting for Approva' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "'  AND S.Merchant IN ('" & _Merchant & "') "

            Else
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                     " FROM ZPP_DEL z INNER JOIN  M01Sales_Order_SAP ON Z.Sales_Order=M01Sales_Order AND M01Line_Item=Z.Line_Item " & _
                    "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                    "WHERE Z.Location in ('N/C Held & waiting for Approva','Aw Finishing') AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' and z.NC_Comment in ('3.AW SHADE COMMENTS','4.AW CUS APP','5.NEED TO FINISH WITH PENDING APP','6.FINISHED WITH PENDING APP','7.SUBMITTED AS ONGOING','8.AW CUS CARE COMMENTS','1.Aw 1st bulk Pilot','2.AW ONGOING PILOT')"

                SQL = "select * from View_Delay_Rpt where Location in ('N/C Held & waiting for Approva','Aw Finishing') and Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' and NC_Comment in ('3.AW SHADE COMMENTS','4.AW CUS APP','5.NEED TO FINISH WITH PENDING APP','6.FINISHED WITH PENDING APP','7.SUBMITTED AS ONGOING','8.AW CUS CARE COMMENTS','1.Aw 1st bulk Pilot','2.AW ONGOING PILOT')"
            End If

            X = X + 1
            _FirstChr = X
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(T01) Then
                worksheet1.Cells(X, 1) = "N/C Held & waiting for approvel"
                worksheet1.Range("A" & X & ":B" & X).MergeCells = True
                worksheet1.Range("A" & X & ":B" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                '  worksheet1.Range("A" & X, "A" & X).Interior.Color = RGB(141, 180, 227)
                worksheet1.Rows(X).Font.size = 10
                worksheet1.Rows(X).Font.name = "Tahoma"
                worksheet1.Rows(X).Font.Bold = True

                _Chart = 97
                For Y = 1 To 19

                    worksheet1.Range(ChrW(_Chart) & X - 1, ChrW(_Chart) & X - 1).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X - 1, ChrW(_Chart) & X - 1).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X - 1, ChrW(_Chart) & X - 1).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X - 1, ChrW(_Chart) & X - 1).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                    _Chart = _Chart + 1
                Next
                X = X + 1
                _FirstChr = X
                i = 0
                For Each DTRow3 As DataRow In T01.Tables(0).Rows
                    Dim _Tollaranz As Double
                    worksheet1.Rows(X).Font.size = 8
                    worksheet1.Rows(X).Font.name = "Tahoma"

                    _Tollaranz = 0
                    _Tollaranz = T01.Tables(0).Rows(i)("M01SO_Qty") * (T01.Tables(0).Rows(i)("M01Cus_Tol_Min") / 100)
                    '   MsgBox(CDbl(T01.Tables(0).Rows(i)("Order_Qty_mtr")))

                    worksheet1.Cells(X, 1) = T01.Tables(0).Rows(i)("Sales_Order")
                    worksheet1.Cells(X, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 2) = T01.Tables(0).Rows(i)("Line_Item")
                    worksheet1.Cells(X, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 3) = T01.Tables(0).Rows(i)("M01Cuatomer_Name")
                    worksheet1.Cells(X, 4) = T01.Tables(0).Rows(i)("Material")
                    worksheet1.Cells(X, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 5) = T01.Tables(0).Rows(i)("Material_Dis")
                    worksheet1.Cells(X, 6) = T01.Tables(0).Rows(i)("No_Dealy_Day")
                    worksheet1.Cells(X, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 7) = T01.Tables(0).Rows(i)("Delivery_Date")
                    worksheet1.Cells(X, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 8) = T01.Tables(0).Rows(i)("Lst_Confirn_Date")
                    worksheet1.Cells(X, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 9) = T01.Tables(0).Rows(i)("No_Day_Same_Opp")
                    worksheet1.Cells(X, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 10) = T01.Tables(0).Rows(i)("Product_Order")
                    worksheet1.Cells(X, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 11) = T01.Tables(0).Rows(i)("Order_Qty_Kg")
                    ' worksheet1.Cells(X, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(X, 11)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(X, 12) = T01.Tables(0).Rows(i)("Order_Qty_mtr")
                    ' worksheet1.Cells(X, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(X, 11)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(X, 13) = T01.Tables(0).Rows(i)("Status")
                    worksheet1.Cells(X, 14) = T01.Tables(0).Rows(i)("Pln_Comment")
                    worksheet1.Cells(X, 15) = T01.Tables(0).Rows(i)("NC_Comment")
                    worksheet1.Cells(X, 16) = T01.Tables(0).Rows(i)("Delay_reason")
                    worksheet1.Cells(X, 17) = T01.Tables(0).Rows(i)("Order_Type")
                    worksheet1.Cells(X, 17).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 18) = T01.Tables(0).Rows(i)("Brandix_OTD_Week")
                    worksheet1.Cells(X, 19) = T01.Tables(0).Rows(i)("Location")
                    worksheet1.Cells(X, 19).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    _Chart = 97
                    For Y = 1 To 19


                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                        _Chart = _Chart + 1
                    Next

                    X = X + 1

                    i = i + 1
                Next

                _Nc1 = X
                worksheet1.Cells(X, 3) = "Sum of N/C Held & waiting for "
                worksheet1.Range("C" & X, "C" & X).Interior.Color = RGB(141, 180, 227)
                worksheet1.Rows(X).Font.size = 10
                worksheet1.Rows(X).Font.name = "Tahoma"
                worksheet1.Rows(X).Font.Bold = True

                worksheet1.Range("K" & (X)).Formula = "=SUM(K" & _FirstChr & ":K" & (X - 1) & ")"
                worksheet1.Cells(X, 11).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet1.Cells(X, 11)
                range1.NumberFormat = "0.00"
                worksheet1.Range("K" & X, "K" & X).Interior.Color = RGB(141, 180, 227)

                worksheet1.Range("l" & (X)).Formula = "=SUM(L" & _FirstChr & ":L" & (X - 1) & ")"
                worksheet1.Cells(X, 12).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet1.Cells(X, 12)
                range1.NumberFormat = "0.00"
                worksheet1.Range("L" & X, "L" & X).Interior.Color = RGB(141, 180, 227)

                _Chart = 97
                For Y = 1 To 19

                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                    _Chart = _Chart + 1
                Next
            End If
            '========================================================================================
            'N/C Held due to dyeing
            If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                      "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                      "WHERE Z.Location='N/C Held due to dyeing' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') AND S.Department IN ('" & _Department & "') AND S.Merchant IN ('" & _Merchant & "') "
            ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                      "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                      "WHERE Z.Location='N/C Held due to dyeing' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') AND S.Department IN ('" & _Department & "')  "
            ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                        "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                        "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                        "WHERE Z.Location='N/C Held due to dyeing' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') AND S.Merchant IN ('" & _Merchant & "')  "
            ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                        "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                        "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                        "WHERE Z.Location='N/C Held due to dyeing' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND S.Department IN ('" & _Department & "') AND S.Merchant IN ('" & _Merchant & "')  "
            ElseIf Trim(txtCustomer.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                                "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                                "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                                "WHERE Z.Location='N/C Held due to dyeing' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') "
            ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                         "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                         "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                         "WHERE Z.Location='N/C Held due to dyeing' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "'  AND S.Department IN ('" & _Department & "')  "
            ElseIf Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                      "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                      "WHERE Z.Location='N/C Held due to dyeing' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "'  AND S.Merchant IN ('" & _Merchant & "') "

            Else
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                     " FROM ZPP_DEL z INNER JOIN  M01Sales_Order_SAP ON Z.Sales_Order=M01Sales_Order AND M01Line_Item=Z.Line_Item " & _
                    "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                    "WHERE Z.Location in ('N/C Held due to dyeing','Aw Finishing','Aw Dyeing') AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' and z.NC_Comment in ('9.Aw pad Pigment','9.AW PAD PIGMENT','10.AW PAD UV','18.Off Shade bulk','15.Batches Tb over Dyed','11.Held in other reason dyeing','16.Stripped tb Over Dyed','14.Held in other reason  others','17.Held due to Trials','18.Off Shade bulk','19.Off shade Sample','20.Off shade Yarn Dye','21.Wet form - TB reprocess','28.Down  Grade')"


                SQL = "select * from View_Delay_Rpt where Location in ('N/C Held due to dyeing','Aw Finishing','Aw Dyeing','Aw prep for Reprocess') and Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' and NC_Comment in ('9.Aw pad Pigment','9.AW PAD PIGMENT','10.AW PAD UV','18.Off Shade bulk','15.Batches Tb over Dyed','11.Held in other reason dyeing','16.Stripped tb Over Dyed','14.Held in other reason  others','17.Held due to Trials','18.Off Shade bulk','19.Off shade Sample','20.Off shade Yarn Dye','21.Wet form - TB reprocess','28.Down  Grade')"
            End If

            X = X + 1
            _FirstChr = X - 1
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(T01) Then
                worksheet1.Cells(X, 1) = "N/C Held due to dyeing"
                worksheet1.Range("A" & X & ":B" & X).MergeCells = True
                worksheet1.Range("A" & X & ":B" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                '  worksheet1.Range("A" & X, "A" & X).Interior.Color = RGB(141, 180, 227)
                worksheet1.Rows(X).Font.size = 10
                worksheet1.Rows(X).Font.name = "Tahoma"
                worksheet1.Rows(X).Font.Bold = True

                _Chart = 97
                For Y = 1 To 19

                    worksheet1.Range(ChrW(_Chart) & X - 1, ChrW(_Chart) & X - 1).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X - 1, ChrW(_Chart) & X - 1).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X - 1, ChrW(_Chart) & X - 1).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X - 1, ChrW(_Chart) & X - 1).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                    _Chart = _Chart + 1
                Next
                X = X + 1
                _FirstChr = X
                i = 0
                For Each DTRow3 As DataRow In T01.Tables(0).Rows
                    Dim _Tollaranz As Double
                    worksheet1.Rows(X).Font.size = 8
                    worksheet1.Rows(X).Font.name = "Tahoma"

                    _Tollaranz = 0
                    _Tollaranz = T01.Tables(0).Rows(i)("M01SO_Qty") * (T01.Tables(0).Rows(i)("M01Cus_Tol_Min") / 100)
                    '   MsgBox(CDbl(T01.Tables(0).Rows(i)("Order_Qty_mtr")))

                    worksheet1.Cells(X, 1) = T01.Tables(0).Rows(i)("Sales_Order")
                    worksheet1.Cells(X, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 2) = T01.Tables(0).Rows(i)("Line_Item")
                    worksheet1.Cells(X, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 3) = T01.Tables(0).Rows(i)("M01Cuatomer_Name")
                    worksheet1.Cells(X, 4) = T01.Tables(0).Rows(i)("Material")
                    worksheet1.Cells(X, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 5) = T01.Tables(0).Rows(i)("Material_Dis")
                    worksheet1.Cells(X, 6) = T01.Tables(0).Rows(i)("No_Dealy_Day")
                    worksheet1.Cells(X, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 7) = T01.Tables(0).Rows(i)("Delivery_Date")
                    worksheet1.Cells(X, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 8) = T01.Tables(0).Rows(i)("Lst_Confirn_Date")
                    worksheet1.Cells(X, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 9) = T01.Tables(0).Rows(i)("No_Day_Same_Opp")
                    worksheet1.Cells(X, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 10) = T01.Tables(0).Rows(i)("Product_Order")
                    worksheet1.Cells(X, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 11) = T01.Tables(0).Rows(i)("Order_Qty_Kg")
                    ' worksheet1.Cells(X, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(X, 11)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(X, 12) = T01.Tables(0).Rows(i)("Order_Qty_mtr")
                    ' worksheet1.Cells(X, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(X, 11)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(X, 13) = T01.Tables(0).Rows(i)("Status")
                    worksheet1.Cells(X, 14) = T01.Tables(0).Rows(i)("Pln_Comment")
                    worksheet1.Cells(X, 15) = T01.Tables(0).Rows(i)("NC_Comment")
                    worksheet1.Cells(X, 16) = T01.Tables(0).Rows(i)("Delay_reason")
                    worksheet1.Cells(X, 17) = T01.Tables(0).Rows(i)("Order_Type")
                    worksheet1.Cells(X, 17).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 18) = T01.Tables(0).Rows(i)("Brandix_OTD_Week")
                    worksheet1.Cells(X, 19) = T01.Tables(0).Rows(i)("Location")
                    worksheet1.Cells(X, 19).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    _Chart = 97
                    For Y = 1 To 19


                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                        _Chart = _Chart + 1
                    Next

                    X = X + 1

                    i = i + 1
                Next

                _NCApp = X
                worksheet1.Cells(X, 3) = "Sum of N/C Held due to dyeing "
                worksheet1.Range("C" & X, "C" & X).Interior.Color = RGB(141, 180, 227)
                worksheet1.Rows(X).Font.size = 10
                worksheet1.Rows(X).Font.name = "Tahoma"
                worksheet1.Rows(X).Font.Bold = True

                worksheet1.Range("K" & (X)).Formula = "=SUM(K" & _FirstChr & ":K" & (X - 1) & ")"
                worksheet1.Cells(X, 11).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet1.Cells(X, 11)
                range1.NumberFormat = "0.00"
                worksheet1.Range("K" & X, "K" & X).Interior.Color = RGB(141, 180, 227)

                worksheet1.Range("l" & (X)).Formula = "=SUM(L" & _FirstChr & ":L" & (X - 1) & ")"
                worksheet1.Cells(X, 12).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet1.Cells(X, 12)
                range1.NumberFormat = "0.00"
                worksheet1.Range("L" & X, "L" & X).Interior.Color = RGB(141, 180, 227)

                _Chart = 97
                For Y = 1 To 19

                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                    _Chart = _Chart + 1
                Next
            End If
            '====================================================================================================
            'N/C Held due to Knitting
            If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                      "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                      "WHERE Z.Location='N/C Held due to Knitting' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') AND S.Department IN ('" & _Department & "') AND S.Merchant IN ('" & _Merchant & "') "
            ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                      "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                      "WHERE Z.Location='N/C Held due to Knitting' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') AND S.Department IN ('" & _Department & "')  "
            ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                        "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                        "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                        "WHERE Z.Location='N/C Held due to Knitting' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') AND S.Merchant IN ('" & _Merchant & "')  "
            ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                        "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                        "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                        "WHERE Z.Location='N/C Held due to Knitting' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND S.Department IN ('" & _Department & "') AND S.Merchant IN ('" & _Merchant & "')  "
            ElseIf Trim(txtCustomer.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                                "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                                "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                                "WHERE Z.Location='N/C Held due to Knitting' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') "
            ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                         "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                         "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                         "WHERE Z.Location='N/C Held due to Knitting' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "'  AND S.Department IN ('" & _Department & "')  "
            ElseIf Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                      "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                      "WHERE Z.Location='N/C Held due to Knitting' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "'  AND S.Merchant IN ('" & _Merchant & "') "

            Else
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                     " FROM ZPP_DEL z INNER JOIN  M01Sales_Order_SAP ON Z.Sales_Order=M01Sales_Order AND M01Line_Item=Z.Line_Item " & _
                    "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                    "WHERE Z.Location='N/C Held due to Knitting' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' and z.NC_Comment='13.Held in other reason Knitting'"


                SQL = "select * from View_Delay_Rpt where Location ='N/C Held due to Knitting' and Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' and NC_Comment ='13.Held in other reason Knitting'"
            End If

           
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(T01) Then
                X = X + 1
                _FirstChr = X
                worksheet1.Cells(X, 1) = "N/C Held due to Knitting"
                worksheet1.Range("A" & X & ":B" & X).MergeCells = True
                worksheet1.Range("A" & X & ":B" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                '  worksheet1.Range("A" & X, "A" & X).Interior.Color = RGB(141, 180, 227)
                worksheet1.Rows(X).Font.size = 10
                worksheet1.Rows(X).Font.name = "Tahoma"
                worksheet1.Rows(X).Font.Bold = True

                _Chart = 97
                For Y = 1 To 19

                    worksheet1.Range(ChrW(_Chart) & X - 1, ChrW(_Chart) & X - 1).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X - 1, ChrW(_Chart) & X - 1).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X - 1, ChrW(_Chart) & X - 1).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X - 1, ChrW(_Chart) & X - 1).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                    _Chart = _Chart + 1
                Next
                X = X + 1
                _FirstChr = X
                i = 0
                For Each DTRow3 As DataRow In T01.Tables(0).Rows
                    Dim _Tollaranz As Double
                    worksheet1.Rows(X).Font.size = 8
                    worksheet1.Rows(X).Font.name = "Tahoma"

                    _Tollaranz = 0
                    _Tollaranz = T01.Tables(0).Rows(i)("M01SO_Qty") * (T01.Tables(0).Rows(i)("M01Cus_Tol_Min") / 100)
                    '   MsgBox(CDbl(T01.Tables(0).Rows(i)("Order_Qty_mtr")))

                    worksheet1.Cells(X, 1) = T01.Tables(0).Rows(i)("Sales_Order")
                    worksheet1.Cells(X, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 2) = T01.Tables(0).Rows(i)("Line_Item")
                    worksheet1.Cells(X, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 3) = T01.Tables(0).Rows(i)("M01Cuatomer_Name")
                    worksheet1.Cells(X, 4) = T01.Tables(0).Rows(i)("Material")
                    worksheet1.Cells(X, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 5) = T01.Tables(0).Rows(i)("Material_Dis")
                    worksheet1.Cells(X, 6) = T01.Tables(0).Rows(i)("No_Dealy_Day")
                    worksheet1.Cells(X, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 7) = T01.Tables(0).Rows(i)("Delivery_Date")
                    worksheet1.Cells(X, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 8) = T01.Tables(0).Rows(i)("Lst_Confirn_Date")
                    worksheet1.Cells(X, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 9) = T01.Tables(0).Rows(i)("No_Day_Same_Opp")
                    worksheet1.Cells(X, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 10) = T01.Tables(0).Rows(i)("Product_Order")
                    worksheet1.Cells(X, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 11) = T01.Tables(0).Rows(i)("Order_Qty_Kg")
                    ' worksheet1.Cells(X, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(X, 11)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(X, 12) = T01.Tables(0).Rows(i)("Order_Qty_mtr")
                    ' worksheet1.Cells(X, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(X, 11)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(X, 13) = T01.Tables(0).Rows(i)("Status")
                    worksheet1.Cells(X, 14) = T01.Tables(0).Rows(i)("Pln_Comment")
                    worksheet1.Cells(X, 15) = T01.Tables(0).Rows(i)("NC_Comment")
                    worksheet1.Cells(X, 16) = T01.Tables(0).Rows(i)("Delay_reason")
                    worksheet1.Cells(X, 17) = T01.Tables(0).Rows(i)("Order_Type")
                    worksheet1.Cells(X, 17).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 18) = T01.Tables(0).Rows(i)("Brandix_OTD_Week")
                    worksheet1.Cells(X, 19) = T01.Tables(0).Rows(i)("Location")
                    worksheet1.Cells(X, 19).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    _Chart = 97
                    For Y = 1 To 19


                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlDot
                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDot
                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlDot
                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDot

                        _Chart = _Chart + 1
                    Next

                    X = X + 1

                    i = i + 1
                Next

                _NC_KNT = X
                worksheet1.Cells(X, 3) = "Sum of N/C Held due to Knitting"
                worksheet1.Range("C" & X, "C" & X).Interior.Color = RGB(141, 180, 227)
                worksheet1.Rows(X).Font.size = 10
                worksheet1.Rows(X).Font.name = "Tahoma"
                worksheet1.Rows(X).Font.Bold = True

                worksheet1.Range("K" & (X)).Formula = "=SUM(K" & _FirstChr & ":K" & (X - 1) & ")"
                worksheet1.Cells(X, 11).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet1.Cells(X, 11)
                range1.NumberFormat = "0.00"
                worksheet1.Range("K" & X, "K" & X).Interior.Color = RGB(141, 180, 227)

                worksheet1.Range("l" & (X)).Formula = "=SUM(L" & _FirstChr & ":L" & (X - 1) & ")"
                worksheet1.Cells(X, 12).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet1.Cells(X, 12)
                range1.NumberFormat = "0.00"
                worksheet1.Range("L" & X, "L" & X).Interior.Color = RGB(141, 180, 227)

                _Chart = 97
                For Y = 1 To 19

                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                    _Chart = _Chart + 1
                Next
            End If
            '====================================================================================================
            'N/C Held due to Finishing
            If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                      "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                      "WHERE Z.Location='N/C Held due to Finishing' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') AND S.Department IN ('" & _Department & "') AND S.Merchant IN ('" & _Merchant & "') "
            ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                      "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                      "WHERE Z.Location='N/C Held due to Finishing' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') AND S.Department IN ('" & _Department & "')  "
            ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                        "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                        "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                        "WHERE Z.Location='N/C Held due to Finishing' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') AND S.Merchant IN ('" & _Merchant & "')  "
            ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                        "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                        "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                        "WHERE Z.Location='N/C Held due to Finishing' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND S.Department IN ('" & _Department & "') AND S.Merchant IN ('" & _Merchant & "')  "
            ElseIf Trim(txtCustomer.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                                "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                                "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                                "WHERE Z.Location='N/C Held due to Finishing' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') "
            ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                         "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                         "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                         "WHERE Z.Location='N/C Held due to Finishing' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "'  AND S.Department IN ('" & _Department & "')  "
            ElseIf Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                      "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                      "WHERE Z.Location='N/C Held due to Finishing' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "'  AND S.Merchant IN ('" & _Merchant & "') "

            Else
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                     " FROM ZPP_DEL z INNER JOIN  M01Sales_Order_SAP ON Z.Sales_Order=M01Sales_Order AND M01Line_Item=Z.Line_Item " & _
                    "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                    "WHERE Z.Location='N/C Held due to Finishing' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' and z.NC_Comment='12.HELD IN OTHER REASON FINISHING'"


                SQL = "select * from View_Delay_Rpt where Location ='N/C Held due to Finishing' and Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' and NC_Comment ='12.HELD IN OTHER REASON FINISHING'"
            End If

          
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(T01) Then
                X = X + 1
                _FirstChr = X

                worksheet1.Cells(X, 1) = "N/C Held due to Finishing"
                worksheet1.Range("A" & X & ":B" & X).MergeCells = True
                worksheet1.Range("A" & X & ":B" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                '  worksheet1.Range("A" & X, "A" & X).Interior.Color = RGB(141, 180, 227)
                worksheet1.Rows(X).Font.size = 10
                worksheet1.Rows(X).Font.name = "Tahoma"
                worksheet1.Rows(X).Font.Bold = True

                _Chart = 97
                For Y = 1 To 19

                    worksheet1.Range(ChrW(_Chart) & X - 1, ChrW(_Chart) & X - 1).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X - 1, ChrW(_Chart) & X - 1).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X - 1, ChrW(_Chart) & X - 1).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X - 1, ChrW(_Chart) & X - 1).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                    _Chart = _Chart + 1
                Next
                X = X + 1
                _FirstChr = X
                i = 0
                For Each DTRow3 As DataRow In T01.Tables(0).Rows
                    Dim _Tollaranz As Double
                    worksheet1.Rows(X).Font.size = 8
                    worksheet1.Rows(X).Font.name = "Tahoma"

                    _Tollaranz = 0
                    _Tollaranz = T01.Tables(0).Rows(i)("M01SO_Qty") * (T01.Tables(0).Rows(i)("M01Cus_Tol_Min") / 100)
                    '   MsgBox(CDbl(T01.Tables(0).Rows(i)("Order_Qty_mtr")))

                    worksheet1.Cells(X, 1) = T01.Tables(0).Rows(i)("Sales_Order")
                    worksheet1.Cells(X, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 2) = T01.Tables(0).Rows(i)("Line_Item")
                    worksheet1.Cells(X, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 3) = T01.Tables(0).Rows(i)("M01Cuatomer_Name")
                    worksheet1.Cells(X, 4) = T01.Tables(0).Rows(i)("Material")
                    worksheet1.Cells(X, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 5) = T01.Tables(0).Rows(i)("Material_Dis")
                    worksheet1.Cells(X, 6) = T01.Tables(0).Rows(i)("No_Dealy_Day")
                    worksheet1.Cells(X, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 7) = T01.Tables(0).Rows(i)("Delivery_Date")
                    worksheet1.Cells(X, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 8) = T01.Tables(0).Rows(i)("Lst_Confirn_Date")
                    worksheet1.Cells(X, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 9) = T01.Tables(0).Rows(i)("No_Day_Same_Opp")
                    worksheet1.Cells(X, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 10) = T01.Tables(0).Rows(i)("Product_Order")
                    worksheet1.Cells(X, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 11) = T01.Tables(0).Rows(i)("Order_Qty_Kg")
                    ' worksheet1.Cells(X, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(X, 11)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(X, 12) = T01.Tables(0).Rows(i)("Order_Qty_mtr")
                    ' worksheet1.Cells(X, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(X, 11)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(X, 13) = T01.Tables(0).Rows(i)("Status")
                    worksheet1.Cells(X, 14) = T01.Tables(0).Rows(i)("Pln_Comment")
                    worksheet1.Cells(X, 15) = T01.Tables(0).Rows(i)("NC_Comment")
                    worksheet1.Cells(X, 16) = T01.Tables(0).Rows(i)("Delay_reason")
                    worksheet1.Cells(X, 17) = T01.Tables(0).Rows(i)("Order_Type")
                    worksheet1.Cells(X, 17).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 18) = T01.Tables(0).Rows(i)("Brandix_OTD_Week")
                    worksheet1.Cells(X, 19) = T01.Tables(0).Rows(i)("Location")
                    worksheet1.Cells(X, 19).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    _Chart = 97
                    For Y = 1 To 19


                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                        _Chart = _Chart + 1
                    Next

                    X = X + 1

                    i = i + 1
                Next

                _NC_FINISHING = X
                worksheet1.Cells(X, 3) = "Sum of N/C Held due to Finishing"
                worksheet1.Range("C" & X, "C" & X).Interior.Color = RGB(141, 180, 227)
                worksheet1.Rows(X).Font.size = 10
                worksheet1.Rows(X).Font.name = "Tahoma"
                worksheet1.Rows(X).Font.Bold = True

                worksheet1.Range("K" & (X)).Formula = "=SUM(K" & _FirstChr & ":K" & (X - 1) & ")"
                worksheet1.Cells(X, 11).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet1.Cells(X, 11)
                range1.NumberFormat = "0.00"
                worksheet1.Range("K" & X, "K" & X).Interior.Color = RGB(141, 180, 227)

                worksheet1.Range("l" & (X)).Formula = "=SUM(L" & _FirstChr & ":L" & (X - 1) & ")"
                worksheet1.Cells(X, 12).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet1.Cells(X, 12)
                range1.NumberFormat = "0.00"
                worksheet1.Range("L" & X, "L" & X).Interior.Color = RGB(141, 180, 227)

                _Chart = 97
                For Y = 1 To 19

                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                    _Chart = _Chart + 1
                Next
            End If
            '====================================================================================================
            'Aw  contamination picking
            If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                      "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                      "WHERE Z.Location='Aw  contamination picking' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') AND S.Department IN ('" & _Department & "') AND S.Merchant IN ('" & _Merchant & "') "
            ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                      "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                      "WHERE Z.Location='Aw  contamination picking' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') AND S.Department IN ('" & _Department & "')  "
            ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                        "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                        "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                        "WHERE Z.Location='Aw  contamination picking' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') AND S.Merchant IN ('" & _Merchant & "')  "
            ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                        "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                        "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                        "WHERE Z.Location='Aw  contamination picking' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND S.Department IN ('" & _Department & "') AND S.Merchant IN ('" & _Merchant & "')  "
            ElseIf Trim(txtCustomer.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                                "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                                "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                                "WHERE Z.Location='Aw  contamination picking' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') "
            ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                         "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                         "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                         "WHERE Z.Location='Aw  contamination picking' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "'  AND S.Department IN ('" & _Department & "')  "
            ElseIf Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                      "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                      "WHERE Z.Location='Aw  contamination picking' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "'  AND S.Merchant IN ('" & _Merchant & "') "

            Else
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                     " FROM ZPP_DEL z INNER JOIN  M01Sales_Order_SAP ON Z.Sales_Order=M01Sales_Order AND M01Line_Item=Z.Line_Item " & _
                    "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                    "WHERE Z.Location='Aw  contamination picking' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' and z.NC_Comment='29.Aw Picking'"

                SQL = "select * from View_Delay_Rpt where Location ='Aw  contamination picking' and Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' and NC_Comment ='29.Aw Picking'"

            End If

            
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(T01) Then
                X = X + 1
                _FirstChr = X

                worksheet1.Cells(X, 1) = "Aw  contamination picking"
                worksheet1.Range("A" & X & ":B" & X).MergeCells = True
                worksheet1.Range("A" & X & ":B" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                '  worksheet1.Range("A" & X, "A" & X).Interior.Color = RGB(141, 180, 227)
                worksheet1.Rows(X).Font.size = 10
                worksheet1.Rows(X).Font.name = "Tahoma"
                worksheet1.Rows(X).Font.Bold = True

                _Chart = 97
                For Y = 1 To 19

                    worksheet1.Range(ChrW(_Chart) & X - 1, ChrW(_Chart) & X - 1).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X - 1, ChrW(_Chart) & X - 1).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X - 1, ChrW(_Chart) & X - 1).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X - 1, ChrW(_Chart) & X - 1).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                    _Chart = _Chart + 1
                Next
                X = X + 1
                _FirstChr = X
                i = 0
                For Each DTRow3 As DataRow In T01.Tables(0).Rows
                    Dim _Tollaranz As Double
                    worksheet1.Rows(X).Font.size = 8
                    worksheet1.Rows(X).Font.name = "Tahoma"

                    _Tollaranz = 0
                    _Tollaranz = T01.Tables(0).Rows(i)("M01SO_Qty") * (T01.Tables(0).Rows(i)("M01Cus_Tol_Min") / 100)
                    '   MsgBox(CDbl(T01.Tables(0).Rows(i)("Order_Qty_mtr")))

                    worksheet1.Cells(X, 1) = T01.Tables(0).Rows(i)("Sales_Order")
                    worksheet1.Cells(X, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 2) = T01.Tables(0).Rows(i)("Line_Item")
                    worksheet1.Cells(X, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 3) = T01.Tables(0).Rows(i)("M01Cuatomer_Name")
                    worksheet1.Cells(X, 4) = T01.Tables(0).Rows(i)("Material")
                    worksheet1.Cells(X, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 5) = T01.Tables(0).Rows(i)("Material_Dis")
                    worksheet1.Cells(X, 6) = T01.Tables(0).Rows(i)("No_Dealy_Day")
                    worksheet1.Cells(X, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 7) = T01.Tables(0).Rows(i)("Delivery_Date")
                    worksheet1.Cells(X, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 8) = T01.Tables(0).Rows(i)("Lst_Confirn_Date")
                    worksheet1.Cells(X, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 9) = T01.Tables(0).Rows(i)("No_Day_Same_Opp")
                    worksheet1.Cells(X, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 10) = T01.Tables(0).Rows(i)("Product_Order")
                    worksheet1.Cells(X, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 11) = T01.Tables(0).Rows(i)("Order_Qty_Kg")
                    ' worksheet1.Cells(X, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(X, 11)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(X, 12) = T01.Tables(0).Rows(i)("Order_Qty_mtr")
                    ' worksheet1.Cells(X, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(X, 11)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(X, 13) = T01.Tables(0).Rows(i)("Status")
                    worksheet1.Cells(X, 14) = T01.Tables(0).Rows(i)("Pln_Comment")
                    worksheet1.Cells(X, 15) = T01.Tables(0).Rows(i)("NC_Comment")
                    worksheet1.Cells(X, 16) = T01.Tables(0).Rows(i)("Delay_reason")
                    worksheet1.Cells(X, 17) = T01.Tables(0).Rows(i)("Order_Type")
                    worksheet1.Cells(X, 17).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 18) = T01.Tables(0).Rows(i)("Brandix_OTD_Week")
                    worksheet1.Cells(X, 19) = T01.Tables(0).Rows(i)("Location")
                    worksheet1.Cells(X, 19).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    _Chart = 97
                    For Y = 1 To 19


                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                        _Chart = _Chart + 1
                    Next

                    X = X + 1

                    i = i + 1
                Next

                _AWCONTAMI = X
                worksheet1.Cells(X, 3) = "Sum of Aw  contamination picking"
                worksheet1.Range("C" & X, "C" & X).Interior.Color = RGB(141, 180, 227)
                worksheet1.Rows(X).Font.size = 10
                worksheet1.Rows(X).Font.name = "Tahoma"
                worksheet1.Rows(X).Font.Bold = True

                worksheet1.Range("K" & (X)).Formula = "=SUM(K" & _FirstChr & ":K" & (X - 1) & ")"
                worksheet1.Cells(X, 11).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet1.Cells(X, 11)
                range1.NumberFormat = "0.00"
                worksheet1.Range("K" & X, "K" & X).Interior.Color = RGB(141, 180, 227)

                worksheet1.Range("l" & (X)).Formula = "=SUM(L" & _FirstChr & ":L" & (X - 1) & ")"
                worksheet1.Cells(X, 12).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet1.Cells(X, 12)
                range1.NumberFormat = "0.00"
                worksheet1.Range("L" & X, "L" & X).Interior.Color = RGB(141, 180, 227)

                _Chart = 97
                For Y = 1 To 19

                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                    _Chart = _Chart + 1
                Next
            End If
            '====================================================================================================
            'AW Print Batches
            If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                      "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                      "WHERE Z.Location='AW Print Batches' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') AND S.Department IN ('" & _Department & "') AND S.Merchant IN ('" & _Merchant & "') "
            ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                      "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                      "WHERE Z.Location='AW Print Batches' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') AND S.Department IN ('" & _Department & "')  "
            ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                        "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                        "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                        "WHERE Z.Location='AW Print Batches' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') AND S.Merchant IN ('" & _Merchant & "')  "
            ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                        "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                        "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                        "WHERE Z.Location='AW Print Batches' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND S.Department IN ('" & _Department & "') AND S.Merchant IN ('" & _Merchant & "')  "
            ElseIf Trim(txtCustomer.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                                "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                                "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                                "WHERE Z.Location='AW Print Batches' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') "
            ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                         "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                         "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                         "WHERE Z.Location='AW Print Batches' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "'  AND S.Department IN ('" & _Department & "')  "
            ElseIf Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                      "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                      "WHERE Z.Location='AW Print Batches' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "'  AND S.Merchant IN ('" & _Merchant & "') "

            Else
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                     " FROM ZPP_DEL z INNER JOIN  M01Sales_Order_SAP ON Z.Sales_Order=M01Sales_Order AND M01Line_Item=Z.Line_Item " & _
                    "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                    "WHERE Z.Location='AW Print Batches' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "'"

                SQL = "select * from View_Delay_Rpt where Location ='AW Print Batches' and Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "'"


            End If

           
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(T01) Then
                X = X + 1
                _FirstChr = X

                worksheet1.Cells(X, 1) = "AW Print Batches"
                worksheet1.Range("A" & X & ":B" & X).MergeCells = True
                worksheet1.Range("A" & X & ":B" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                '  worksheet1.Range("A" & X, "A" & X).Interior.Color = RGB(141, 180, 227)
                worksheet1.Rows(X).Font.size = 10
                worksheet1.Rows(X).Font.name = "Tahoma"
                worksheet1.Rows(X).Font.Bold = True

                _Chart = 97
                For Y = 1 To 19

                    worksheet1.Range(ChrW(_Chart) & X - 1, ChrW(_Chart) & X - 1).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X - 1, ChrW(_Chart) & X - 1).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X - 1, ChrW(_Chart) & X - 1).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X - 1, ChrW(_Chart) & X - 1).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                    _Chart = _Chart + 1
                Next
                X = X + 1
                _FirstChr = X
                i = 0
                For Each DTRow3 As DataRow In T01.Tables(0).Rows
                    Dim _Tollaranz As Double
                    worksheet1.Rows(X).Font.size = 8
                    worksheet1.Rows(X).Font.name = "Tahoma"

                    _Tollaranz = 0
                    _Tollaranz = T01.Tables(0).Rows(i)("M01SO_Qty") * (T01.Tables(0).Rows(i)("M01Cus_Tol_Min") / 100)
                    '   MsgBox(CDbl(T01.Tables(0).Rows(i)("Order_Qty_mtr")))

                    worksheet1.Cells(X, 1) = T01.Tables(0).Rows(i)("Sales_Order")
                    worksheet1.Cells(X, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 2) = T01.Tables(0).Rows(i)("Line_Item")
                    worksheet1.Cells(X, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 3) = T01.Tables(0).Rows(i)("M01Cuatomer_Name")
                    worksheet1.Cells(X, 4) = T01.Tables(0).Rows(i)("Material")
                    worksheet1.Cells(X, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 5) = T01.Tables(0).Rows(i)("Material_Dis")
                    worksheet1.Cells(X, 6) = T01.Tables(0).Rows(i)("No_Dealy_Day")
                    worksheet1.Cells(X, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 7) = T01.Tables(0).Rows(i)("Delivery_Date")
                    worksheet1.Cells(X, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 8) = T01.Tables(0).Rows(i)("Lst_Confirn_Date")
                    worksheet1.Cells(X, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 9) = T01.Tables(0).Rows(i)("No_Day_Same_Opp")
                    worksheet1.Cells(X, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 10) = T01.Tables(0).Rows(i)("Product_Order")
                    worksheet1.Cells(X, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 11) = T01.Tables(0).Rows(i)("Order_Qty_Kg")
                    ' worksheet1.Cells(X, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(X, 11)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(X, 12) = T01.Tables(0).Rows(i)("Order_Qty_mtr")
                    ' worksheet1.Cells(X, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(X, 11)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(X, 13) = T01.Tables(0).Rows(i)("Status")
                    worksheet1.Cells(X, 14) = T01.Tables(0).Rows(i)("Pln_Comment")
                    worksheet1.Cells(X, 15) = T01.Tables(0).Rows(i)("NC_Comment")
                    worksheet1.Cells(X, 16) = T01.Tables(0).Rows(i)("Delay_reason")
                    worksheet1.Cells(X, 17) = T01.Tables(0).Rows(i)("Order_Type")
                    worksheet1.Cells(X, 17).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 18) = T01.Tables(0).Rows(i)("Brandix_OTD_Week")
                    worksheet1.Cells(X, 19) = T01.Tables(0).Rows(i)("Location")
                    worksheet1.Cells(X, 19).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    _Chart = 97
                    For Y = 1 To 19


                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                        _Chart = _Chart + 1
                    Next

                    X = X + 1

                    i = i + 1
                Next

                _AWPrint = X
                worksheet1.Cells(X, 3) = "Sum of AW Print Batches"
                worksheet1.Range("C" & X, "C" & X).Interior.Color = RGB(141, 180, 227)
                worksheet1.Rows(X).Font.size = 10
                worksheet1.Rows(X).Font.name = "Tahoma"
                worksheet1.Rows(X).Font.Bold = True

                worksheet1.Range("K" & (X)).Formula = "=SUM(K" & _FirstChr & ":K" & (X - 1) & ")"
                worksheet1.Cells(X, 11).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet1.Cells(X, 11)
                range1.NumberFormat = "0.00"
                worksheet1.Range("K" & X, "K" & X).Interior.Color = RGB(141, 180, 227)

                worksheet1.Range("l" & (X)).Formula = "=SUM(L" & _FirstChr & ":L" & (X - 1) & ")"
                worksheet1.Cells(X, 12).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet1.Cells(X, 12)
                range1.NumberFormat = "0.00"
                worksheet1.Range("L" & X, "L" & X).Interior.Color = RGB(141, 180, 227)

                _Chart = 97
                For Y = 1 To 19

                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                    _Chart = _Chart + 1
                Next
            End If
            '===================================================================================================
            'Aw prep for Reprocess
            If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                      "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                      "WHERE Z.Location='Aw prep for Reprocess' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') AND S.Department IN ('" & _Department & "') AND S.Merchant IN ('" & _Merchant & "') "
            ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                      "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                      "WHERE Z.Location='Aw prep for Reprocess' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') AND S.Department IN ('" & _Department & "')  "
            ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                        "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                        "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                        "WHERE Z.Location='Aw prep for Reprocess' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') AND S.Merchant IN ('" & _Merchant & "')  "
            ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                        "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                        "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                        "WHERE Z.Location='Aw prep for Reprocess' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND S.Department IN ('" & _Department & "') AND S.Merchant IN ('" & _Merchant & "')  "
            ElseIf Trim(txtCustomer.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                                "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                                "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                                "WHERE Z.Location='Aw prep for Reprocess' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') "
            ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                         "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                         "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                         "WHERE Z.Location='Aw prep for Reprocess' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "'  AND S.Department IN ('" & _Department & "')  "
            ElseIf Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                      "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                      "WHERE Z.Location='Aw prep for Reprocess' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "'  AND S.Merchant IN ('" & _Merchant & "') "

            Else
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                     " FROM ZPP_DEL z INNER JOIN  M01Sales_Order_SAP ON Z.Sales_Order=M01Sales_Order AND M01Line_Item=Z.Line_Item " & _
                    "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                    "WHERE Z.Location='Aw prep for Reprocess' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "'"


                SQL = "select * from View_Delay_Rpt where Location ='Aw prep for Reprocess' and Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' and NC_Comment=''"
            End If

            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(T01) Then

                X = X + 1
                _FirstChr = X

                worksheet1.Cells(X, 1) = "Aw prep for Reprocess"
                worksheet1.Range("A" & X & ":B" & X).MergeCells = True
                worksheet1.Range("A" & X & ":B" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                '  worksheet1.Range("A" & X, "A" & X).Interior.Color = RGB(141, 180, 227)
                worksheet1.Rows(X).Font.size = 10
                worksheet1.Rows(X).Font.name = "Tahoma"
                worksheet1.Rows(X).Font.Bold = True

                _Chart = 97
                For Y = 1 To 19

                    worksheet1.Range(ChrW(_Chart) & X - 1, ChrW(_Chart) & X - 1).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X - 1, ChrW(_Chart) & X - 1).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X - 1, ChrW(_Chart) & X - 1).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X - 1, ChrW(_Chart) & X - 1).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                    _Chart = _Chart + 1
                Next
                X = X + 1
                _FirstChr = X
                i = 0
                For Each DTRow3 As DataRow In T01.Tables(0).Rows
                    Dim _Tollaranz As Double
                    worksheet1.Rows(X).Font.size = 8
                    worksheet1.Rows(X).Font.name = "Tahoma"

                    _Tollaranz = 0
                    _Tollaranz = T01.Tables(0).Rows(i)("M01SO_Qty") * (T01.Tables(0).Rows(i)("M01Cus_Tol_Min") / 100)
                    '   MsgBox(CDbl(T01.Tables(0).Rows(i)("Order_Qty_mtr")))

                    worksheet1.Cells(X, 1) = T01.Tables(0).Rows(i)("Sales_Order")
                    worksheet1.Cells(X, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 2) = T01.Tables(0).Rows(i)("Line_Item")
                    worksheet1.Cells(X, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 3) = T01.Tables(0).Rows(i)("M01Cuatomer_Name")
                    worksheet1.Cells(X, 4) = T01.Tables(0).Rows(i)("Material")
                    worksheet1.Cells(X, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 5) = T01.Tables(0).Rows(i)("Material_Dis")
                    worksheet1.Cells(X, 6) = T01.Tables(0).Rows(i)("No_Dealy_Day")
                    worksheet1.Cells(X, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 7) = T01.Tables(0).Rows(i)("Delivery_Date")
                    worksheet1.Cells(X, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 8) = T01.Tables(0).Rows(i)("Lst_Confirn_Date")
                    worksheet1.Cells(X, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 9) = T01.Tables(0).Rows(i)("No_Day_Same_Opp")
                    worksheet1.Cells(X, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 10) = T01.Tables(0).Rows(i)("Product_Order")
                    worksheet1.Cells(X, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 11) = T01.Tables(0).Rows(i)("Order_Qty_Kg")
                    ' worksheet1.Cells(X, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(X, 11)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(X, 12) = T01.Tables(0).Rows(i)("Order_Qty_mtr")
                    ' worksheet1.Cells(X, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(X, 11)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(X, 13) = T01.Tables(0).Rows(i)("Status")
                    worksheet1.Cells(X, 14) = T01.Tables(0).Rows(i)("Pln_Comment")
                    worksheet1.Cells(X, 15) = T01.Tables(0).Rows(i)("NC_Comment")
                    worksheet1.Cells(X, 16) = T01.Tables(0).Rows(i)("Delay_reason")
                    worksheet1.Cells(X, 17) = T01.Tables(0).Rows(i)("Order_Type")
                    worksheet1.Cells(X, 17).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 18) = T01.Tables(0).Rows(i)("Brandix_OTD_Week")
                    worksheet1.Cells(X, 19) = T01.Tables(0).Rows(i)("Location")
                    worksheet1.Cells(X, 19).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    _Chart = 97
                    For Y = 1 To 19


                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                        _Chart = _Chart + 1
                    Next

                    X = X + 1

                    i = i + 1
                Next

                _Awprep = X
                worksheet1.Cells(X, 3) = "Sum of Aw prep for Reprocess "
                worksheet1.Range("C" & X, "C" & X).Interior.Color = RGB(141, 180, 227)
                worksheet1.Rows(X).Font.size = 10
                worksheet1.Rows(X).Font.name = "Tahoma"
                worksheet1.Rows(X).Font.Bold = True

                worksheet1.Range("K" & (X)).Formula = "=SUM(K" & _FirstChr & ":K" & (X - 1) & ")"
                worksheet1.Cells(X, 11).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet1.Cells(X, 11)
                range1.NumberFormat = "0.00"
                worksheet1.Range("K" & X, "K" & X).Interior.Color = RGB(141, 180, 227)

                worksheet1.Range("l" & (X)).Formula = "=SUM(L" & _FirstChr & ":L" & (X - 1) & ")"
                worksheet1.Cells(X, 12).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet1.Cells(X, 12)
                range1.NumberFormat = "0.00"
                worksheet1.Range("L" & X, "L" & X).Interior.Color = RGB(141, 180, 227)

                _Chart = 97
                For Y = 1 To 19

                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                    _Chart = _Chart + 1
                Next
            End If
            '===========================================================================
            'Aw Finishing
            'Remove s.Del_Date=z.Delivery_Date request by Amila on 3/3/2016
            If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                      "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item  " & _
                      "WHERE Z.Location='Aw Finishing' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') AND S.Department IN ('" & _Department & "') AND S.Merchant IN ('" & _Merchant & "') "
            ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                      "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item  " & _
                      "WHERE Z.Location='Aw Finishing' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') AND S.Department IN ('" & _Department & "')  "
            ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                        "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                        "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item  " & _
                        "WHERE Z.Location='Aw Finishing' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') AND S.Merchant IN ('" & _Merchant & "')  "
            ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                        "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                        "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item  " & _
                        "WHERE Z.Location='Aw Finishing' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND S.Department IN ('" & _Department & "') AND S.Merchant IN ('" & _Merchant & "')  "
            ElseIf Trim(txtCustomer.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                                "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                                "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item " & _
                                "WHERE Z.Location='Aw Finishing' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') "
            ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                         "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                         "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item  " & _
                         "WHERE Z.Location='Aw Finishing' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "'  AND S.Department IN ('" & _Department & "')  "
            ElseIf Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                      "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item  " & _
                      "WHERE Z.Location='Aw Finishing' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "'  AND S.Merchant IN ('" & _Merchant & "') "

            Else
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                     " FROM ZPP_DEL z INNER JOIN  M01Sales_Order_SAP ON Z.Sales_Order=M01Sales_Order AND M01Line_Item=Z.Line_Item " & _
                    "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item  " & _
                    "WHERE Z.Location='Aw Finishing' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' and z.NC_Comment not in ('22.RESCHEDULE - RE DYE','23.RESCHEDULE – WASH','24.RESCHEDULE – STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RE SCHEDULE – SAMPLE','9.Aw pad Pigment','9.AW PAD PIGMENT','10.AW PAD UV','18.Off Shade bulk','15.Batches Tb over Dyed','11.Held in other reason dyeing','16.Stripped tb Over Dyed','14.Held in other reason  others','17.Held due to Trials','18.Off Shade bulk','19.Off shade Sample','20.Off shade Yarn Dye','21.Wet form - TB reprocess','28.Down  Grade','3.AW SHADE COMMENTS','4.AW CUS APP','5.NEED TO FINISH WITH PENDING APP','6.FINISHED WITH PENDING APP','7.SUBMITTED AS ONGOING','8.AW CUS CARE COMMENTS','1.Aw 1st bulk Pilot','2.AW ONGOING PILOT')"

                SQL = "select * from View_Delay_Rpt where Location ='Aw Finishing' and Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' and NC_Comment not in ('22.RESCHEDULE - RE DYE','23.RESCHEDULE – WASH','24.RESCHEDULE – STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RE SCHEDULE – SAMPLE','9.Aw pad Pigment','9.AW PAD PIGMENT','10.AW PAD UV','18.Off Shade bulk','15.Batches Tb over Dyed','11.Held in other reason dyeing','16.Stripped tb Over Dyed','14.Held in other reason  others','17.Held due to Trials','18.Off Shade bulk','19.Off shade Sample','20.Off shade Yarn Dye','21.Wet form - TB reprocess','28.Down  Grade','3.AW SHADE COMMENTS','4.AW CUS APP','5.NEED TO FINISH WITH PENDING APP','6.FINISHED WITH PENDING APP','7.SUBMITTED AS ONGOING','8.AW CUS CARE COMMENTS','1.Aw 1st bulk Pilot','2.AW ONGOING PILOT')"

            End If

            
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(T01) Then
                X = X + 1
                _FirstChr = X
                worksheet1.Cells(X, 1) = "Aw Finishing"
                worksheet1.Range("A" & X & ":B" & X).MergeCells = True
                worksheet1.Range("A" & X & ":B" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                '  worksheet1.Range("A" & X, "A" & X).Interior.Color = RGB(141, 180, 227)
                worksheet1.Rows(X).Font.size = 10
                worksheet1.Rows(X).Font.name = "Tahoma"
                worksheet1.Rows(X).Font.Bold = True

                _Chart = 97
                For Y = 1 To 19

                    worksheet1.Range(ChrW(_Chart) & X - 1, ChrW(_Chart) & X - 1).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X - 1, ChrW(_Chart) & X - 1).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X - 1, ChrW(_Chart) & X - 1).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X - 1, ChrW(_Chart) & X - 1).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                    _Chart = _Chart + 1
                Next
                X = X + 1
                _FirstChr = X
                i = 0
                For Each DTRow3 As DataRow In T01.Tables(0).Rows
                    Dim _Tollaranz As Double
                    worksheet1.Rows(X).Font.size = 8
                    worksheet1.Rows(X).Font.name = "Tahoma"

                    _Tollaranz = 0
                    _Tollaranz = T01.Tables(0).Rows(i)("M01SO_Qty") * (T01.Tables(0).Rows(i)("M01Cus_Tol_Min") / 100)
                    '   MsgBox(CDbl(T01.Tables(0).Rows(i)("Order_Qty_mtr")))

                    worksheet1.Cells(X, 1) = T01.Tables(0).Rows(i)("Sales_Order")
                    worksheet1.Cells(X, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 2) = T01.Tables(0).Rows(i)("Line_Item")
                    worksheet1.Cells(X, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 3) = T01.Tables(0).Rows(i)("M01Cuatomer_Name")
                    worksheet1.Cells(X, 4) = T01.Tables(0).Rows(i)("Material")
                    worksheet1.Cells(X, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 5) = T01.Tables(0).Rows(i)("Material_Dis")
                    worksheet1.Cells(X, 6) = T01.Tables(0).Rows(i)("No_Dealy_Day")
                    worksheet1.Cells(X, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 7) = T01.Tables(0).Rows(i)("Delivery_Date")
                    worksheet1.Cells(X, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 8) = T01.Tables(0).Rows(i)("Lst_Confirn_Date")
                    worksheet1.Cells(X, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 9) = T01.Tables(0).Rows(i)("No_Day_Same_Opp")
                    worksheet1.Cells(X, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 10) = T01.Tables(0).Rows(i)("Product_Order")
                    worksheet1.Cells(X, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 11) = T01.Tables(0).Rows(i)("Order_Qty_Kg")
                    ' worksheet1.Cells(X, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(X, 11)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(X, 12) = T01.Tables(0).Rows(i)("Order_Qty_mtr")
                    ' worksheet1.Cells(X, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(X, 11)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(X, 13) = T01.Tables(0).Rows(i)("Status")
                    worksheet1.Cells(X, 14) = T01.Tables(0).Rows(i)("Pln_Comment")
                    worksheet1.Cells(X, 15) = T01.Tables(0).Rows(i)("NC_Comment")
                    worksheet1.Cells(X, 16) = T01.Tables(0).Rows(i)("Delay_reason")
                    worksheet1.Cells(X, 17) = T01.Tables(0).Rows(i)("Order_Type")
                    worksheet1.Cells(X, 17).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 18) = T01.Tables(0).Rows(i)("Brandix_OTD_Week")
                    worksheet1.Cells(X, 19) = T01.Tables(0).Rows(i)("Location")
                    worksheet1.Cells(X, 19).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    _Chart = 97
                    For Y = 1 To 19


                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                        _Chart = _Chart + 1
                    Next

                    X = X + 1

                    i = i + 1
                Next

                _Aw_Finishing = X
                worksheet1.Cells(X, 3) = "Sum of Aw Finishing "
                worksheet1.Range("C" & X, "C" & X).Interior.Color = RGB(141, 180, 227)
                worksheet1.Rows(X).Font.size = 10
                worksheet1.Rows(X).Font.name = "Tahoma"
                worksheet1.Rows(X).Font.Bold = True

                worksheet1.Range("K" & (X)).Formula = "=SUM(K" & _FirstChr & ":K" & (X - 1) & ")"
                worksheet1.Cells(X, 11).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet1.Cells(X, 11)
                range1.NumberFormat = "0.00"
                worksheet1.Range("K" & X, "K" & X).Interior.Color = RGB(141, 180, 227)

                worksheet1.Range("l" & (X)).Formula = "=SUM(L" & _FirstChr & ":L" & (X - 1) & ")"
                worksheet1.Cells(X, 12).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet1.Cells(X, 12)
                range1.NumberFormat = "0.00"
                worksheet1.Range("L" & X, "L" & X).Interior.Color = RGB(141, 180, 227)

                _Chart = 97
                For Y = 1 To 19

                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                    _Chart = _Chart + 1
                Next
            End If
            '==========================================================================================
            'total Aw Exam
            'Remove s.Del_Date=z.Delivery_Date request by Amila
            If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                      "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item  " & _
                      "WHERE Z.Location='Total Aw Exam' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') AND S.Department IN ('" & _Department & "') AND S.Merchant IN ('" & _Merchant & "') "
            ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                      "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item  " & _
                      "WHERE Z.Location='Total Aw Exam' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') AND S.Department IN ('" & _Department & "')  "
            ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                        "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                        "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item  " & _
                        "WHERE Z.Location='Total Aw Exam' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') AND S.Merchant IN ('" & _Merchant & "')  "
            ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                        "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                        "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item  " & _
                        "WHERE Z.Location='Total Aw Exam' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND S.Department IN ('" & _Department & "') AND S.Merchant IN ('" & _Merchant & "')  "
            ElseIf Trim(txtCustomer.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                                "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                                "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item  " & _
                                "WHERE Z.Location='Total Aw Exam' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') "
            ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                         "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                         "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item  " & _
                         "WHERE Z.Location='Total Aw Exam' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "'  AND S.Department IN ('" & _Department & "')  "
            ElseIf Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                      "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item  " & _
                      "WHERE Z.Location='Total Aw Exam' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "'  AND S.Merchant IN ('" & _Merchant & "') "

            Else
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                     " FROM ZPP_DEL z INNER JOIN  M01Sales_Order_SAP ON Z.Sales_Order=M01Sales_Order AND M01Line_Item=Z.Line_Item " & _
                    "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item  " & _
                    "WHERE Z.Location='Total Aw Exam' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "'"


                SQL = "select * from View_Delay_Rpt where Location ='Total Aw Exam' and Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' "
            End If

            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(T01) Then

                X = X + 1
                _FirstChr = X
                worksheet1.Cells(X, 1) = "Aw Exam"
                worksheet1.Range("A" & X & ":B" & X).MergeCells = True
                worksheet1.Range("A" & X & ":B" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                '  worksheet1.Range("A" & X, "A" & X).Interior.Color = RGB(141, 180, 227)
                worksheet1.Rows(X).Font.size = 10
                worksheet1.Rows(X).Font.name = "Tahoma"
                worksheet1.Rows(X).Font.Bold = True

                _Chart = 97
                For Y = 1 To 19

                    worksheet1.Range(ChrW(_Chart) & X - 1, ChrW(_Chart) & X - 1).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X - 1, ChrW(_Chart) & X - 1).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X - 1, ChrW(_Chart) & X - 1).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X - 1, ChrW(_Chart) & X - 1).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                    _Chart = _Chart + 1
                Next
                X = X + 1
                _FirstChr = X
                i = 0
                For Each DTRow3 As DataRow In T01.Tables(0).Rows
                    Dim _Tollaranz As Double
                    worksheet1.Rows(X).Font.size = 8
                    worksheet1.Rows(X).Font.name = "Tahoma"

                    _Tollaranz = 0
                    _Tollaranz = T01.Tables(0).Rows(i)("M01SO_Qty") * (T01.Tables(0).Rows(i)("M01Cus_Tol_Min") / 100)
                    '   MsgBox(CDbl(T01.Tables(0).Rows(i)("Order_Qty_mtr")))

                    worksheet1.Cells(X, 1) = T01.Tables(0).Rows(i)("Sales_Order")
                    worksheet1.Cells(X, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 2) = T01.Tables(0).Rows(i)("Line_Item")
                    worksheet1.Cells(X, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 3) = T01.Tables(0).Rows(i)("M01Cuatomer_Name")
                    worksheet1.Cells(X, 4) = T01.Tables(0).Rows(i)("Material")
                    worksheet1.Cells(X, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 5) = T01.Tables(0).Rows(i)("Material_Dis")
                    worksheet1.Cells(X, 6) = T01.Tables(0).Rows(i)("No_Dealy_Day")
                    worksheet1.Cells(X, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 7) = T01.Tables(0).Rows(i)("Delivery_Date")
                    worksheet1.Cells(X, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 8) = T01.Tables(0).Rows(i)("Lst_Confirn_Date")
                    worksheet1.Cells(X, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 9) = T01.Tables(0).Rows(i)("No_Day_Same_Opp")
                    worksheet1.Cells(X, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 10) = T01.Tables(0).Rows(i)("Product_Order")
                    worksheet1.Cells(X, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 11) = T01.Tables(0).Rows(i)("Order_Qty_Kg")
                    ' worksheet1.Cells(X, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(X, 11)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(X, 12) = T01.Tables(0).Rows(i)("Order_Qty_mtr")
                    ' worksheet1.Cells(X, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(X, 11)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(X, 13) = T01.Tables(0).Rows(i)("Status")
                    worksheet1.Cells(X, 14) = T01.Tables(0).Rows(i)("Pln_Comment")
                    worksheet1.Cells(X, 15) = T01.Tables(0).Rows(i)("NC_Comment")
                    worksheet1.Cells(X, 16) = T01.Tables(0).Rows(i)("Delay_reason")
                    worksheet1.Cells(X, 17) = T01.Tables(0).Rows(i)("Order_Type")
                    worksheet1.Cells(X, 17).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 18) = T01.Tables(0).Rows(i)("Brandix_OTD_Week")
                    worksheet1.Cells(X, 19) = T01.Tables(0).Rows(i)("Location")
                    worksheet1.Cells(X, 19).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    _Chart = 97
                    For Y = 1 To 19


                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                        _Chart = _Chart + 1
                    Next

                    X = X + 1

                    i = i + 1
                Next

                _Aw_Exam = X
                worksheet1.Cells(X, 3) = "Sum of Aw Exam "
                worksheet1.Range("C" & X, "C" & X).Interior.Color = RGB(141, 180, 227)
                worksheet1.Rows(X).Font.size = 10
                worksheet1.Rows(X).Font.name = "Tahoma"
                worksheet1.Rows(X).Font.Bold = True

                worksheet1.Range("K" & (X)).Formula = "=SUM(K" & _FirstChr & ":K" & (X - 1) & ")"
                worksheet1.Cells(X, 11).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet1.Cells(X, 11)
                range1.NumberFormat = "0.00"
                worksheet1.Range("K" & X, "K" & X).Interior.Color = RGB(141, 180, 227)

                worksheet1.Range("l" & (X)).Formula = "=SUM(L" & _FirstChr & ":L" & (X - 1) & ")"
                worksheet1.Cells(X, 12).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet1.Cells(X, 12)
                range1.NumberFormat = "0.00"
                worksheet1.Range("L" & X, "L" & X).Interior.Color = RGB(141, 180, 227)

                _Chart = 97
                For Y = 1 To 19

                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                    _Chart = _Chart + 1
                Next
            End If
            '========================================================================
            'Blocked 2065
            If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                      "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                      "WHERE Z.Location='Blocked 2065' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') AND S.Department IN ('" & _Department & "') AND S.Merchant IN ('" & _Merchant & "') "
            ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                      "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                      "WHERE Z.Location='Blocked 2065' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') AND S.Department IN ('" & _Department & "')  "
            ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                        "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                        "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                        "WHERE Z.Location='Blocked 2065' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') AND S.Merchant IN ('" & _Merchant & "')  "
            ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                        "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                        "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                        "WHERE Z.Location='Blocked 2065' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND S.Department IN ('" & _Department & "') AND S.Merchant IN ('" & _Merchant & "')  "
            ElseIf Trim(txtCustomer.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                                "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                                "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                                "WHERE Z.Location='Blocked 2065' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND S.Customer IN ('" & _Customer & "') "
            ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                         "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                         "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                         "WHERE Z.Location='Blocked 2065' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "'  AND S.Department IN ('" & _Department & "')  "
            ElseIf Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                      "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                      "WHERE Z.Location='Blocked 2065' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "'  AND S.Merchant IN ('" & _Merchant & "') "

            Else
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                     " FROM ZPP_DEL z INNER JOIN  M01Sales_Order_SAP ON Z.Sales_Order=M01Sales_Order AND M01Line_Item=Z.Line_Item " & _
                    "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                    "WHERE Z.Location='Blocked 2065' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "'"


                SQL = "select * from View_Delay_Rpt where Location ='Blocked 2065' and Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' "
            End If

           
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(T01) Then
                X = X + 1
                _FirstChr = X

                worksheet1.Cells(X, 1) = "Blocked 2065"
                worksheet1.Range("A" & X & ":B" & X).MergeCells = True
                worksheet1.Range("A" & X & ":B" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                '  worksheet1.Range("A" & X, "A" & X).Interior.Color = RGB(141, 180, 227)
                worksheet1.Rows(X).Font.size = 10
                worksheet1.Rows(X).Font.name = "Tahoma"
                worksheet1.Rows(X).Font.Bold = True

                _Chart = 97
                For Y = 1 To 19

                    worksheet1.Range(ChrW(_Chart) & X - 1, ChrW(_Chart) & X - 1).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X - 1, ChrW(_Chart) & X - 1).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X - 1, ChrW(_Chart) & X - 1).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X - 1, ChrW(_Chart) & X - 1).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                    _Chart = _Chart + 1
                Next
                X = X + 1
                _FirstChr = X
                i = 0
                For Each DTRow3 As DataRow In T01.Tables(0).Rows
                    Dim _Tollaranz As Double
                    worksheet1.Rows(X).Font.size = 8
                    worksheet1.Rows(X).Font.name = "Tahoma"

                    _Tollaranz = 0
                    _Tollaranz = T01.Tables(0).Rows(i)("M01SO_Qty") * (T01.Tables(0).Rows(i)("M01Cus_Tol_Min") / 100)
                    '   MsgBox(CDbl(T01.Tables(0).Rows(i)("Order_Qty_mtr")))

                    worksheet1.Cells(X, 1) = T01.Tables(0).Rows(i)("Sales_Order")
                    worksheet1.Cells(X, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 2) = T01.Tables(0).Rows(i)("Line_Item")
                    worksheet1.Cells(X, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 3) = T01.Tables(0).Rows(i)("M01Cuatomer_Name")
                    worksheet1.Cells(X, 4) = T01.Tables(0).Rows(i)("Material")
                    worksheet1.Cells(X, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 5) = T01.Tables(0).Rows(i)("Material_Dis")
                    worksheet1.Cells(X, 6) = T01.Tables(0).Rows(i)("No_Dealy_Day")
                    worksheet1.Cells(X, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 7) = T01.Tables(0).Rows(i)("Delivery_Date")
                    worksheet1.Cells(X, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 8) = T01.Tables(0).Rows(i)("Lst_Confirn_Date")
                    worksheet1.Cells(X, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 9) = T01.Tables(0).Rows(i)("No_Day_Same_Opp")
                    worksheet1.Cells(X, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 10) = T01.Tables(0).Rows(i)("Product_Order")
                    worksheet1.Cells(X, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 11) = T01.Tables(0).Rows(i)("Order_Qty_Kg")
                    ' worksheet1.Cells(X, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(X, 11)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(X, 12) = T01.Tables(0).Rows(i)("Order_Qty_mtr")
                    ' worksheet1.Cells(X, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(X, 11)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(X, 13) = T01.Tables(0).Rows(i)("Status")
                    worksheet1.Cells(X, 14) = T01.Tables(0).Rows(i)("Pln_Comment")
                    worksheet1.Cells(X, 15) = T01.Tables(0).Rows(i)("NC_Comment")
                    worksheet1.Cells(X, 16) = T01.Tables(0).Rows(i)("Delay_reason")
                    worksheet1.Cells(X, 17) = T01.Tables(0).Rows(i)("Order_Type")
                    worksheet1.Cells(X, 17).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 18) = T01.Tables(0).Rows(i)("Brandix_OTD_Week")
                    worksheet1.Cells(X, 19) = T01.Tables(0).Rows(i)("Location")
                    worksheet1.Cells(X, 19).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    _Chart = 97
                    For Y = 1 To 19


                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                        _Chart = _Chart + 1
                    Next

                    X = X + 1

                    i = i + 1
                Next

                _2065 = X
                worksheet1.Cells(X, 3) = "Sum of Blocked 2065 "
                worksheet1.Range("C" & X, "C" & X).Interior.Color = RGB(141, 180, 227)
                worksheet1.Rows(X).Font.size = 10
                worksheet1.Rows(X).Font.name = "Tahoma"
                worksheet1.Rows(X).Font.Bold = True

                worksheet1.Range("K" & (X)).Formula = "=SUM(K" & _FirstChr & ":K" & (X - 1) & ")"
                worksheet1.Cells(X, 11).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet1.Cells(X, 11)
                range1.NumberFormat = "0.00"
                worksheet1.Range("K" & X, "K" & X).Interior.Color = RGB(141, 180, 227)

                worksheet1.Range("l" & (X)).Formula = "=SUM(L" & _FirstChr & ":L" & (X - 1) & ")"
                worksheet1.Cells(X, 12).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet1.Cells(X, 12)
                range1.NumberFormat = "0.00"
                worksheet1.Range("L" & X, "L" & X).Interior.Color = RGB(141, 180, 227)

                _Chart = 97
                For Y = 1 To 19

                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                    _Chart = _Chart + 1
                Next
            End If
            '==========================================================================
            '2062
            If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                      "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                      "WHERE Z.Location='Blocked 2062' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') AND S.Department IN ('" & _Department & "') AND S.Merchant IN ('" & _Merchant & "') "
            ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                      "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                      "WHERE Z.Location='Blocked 2062' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') AND S.Department IN ('" & _Department & "')  "
            ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                        "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                        "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                        "WHERE Z.Location='Blocked 2062' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') AND S.Merchant IN ('" & _Merchant & "')  "
            ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                        "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                        "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                        "WHERE Z.Location='Blocked 2062' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND S.Department IN ('" & _Department & "') AND S.Merchant IN ('" & _Merchant & "')  "
            ElseIf Trim(txtCustomer.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                                "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                                "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                                "WHERE Z.Location='Blocked 2062' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') "
            ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                         "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                         "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                         "WHERE Z.Location='Blocked 2062' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "'  AND S.Department IN ('" & _Department & "')  "
            ElseIf Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                      "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                      "WHERE Z.Location='Blocked 2062' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "'  AND S.Merchant IN ('" & _Merchant & "') "

            Else
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                     " FROM ZPP_DEL z INNER JOIN  M01Sales_Order_SAP ON Z.Sales_Order=M01Sales_Order AND M01Line_Item=Z.Line_Item " & _
                    "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                    "WHERE Z.Location='Blocked 2062' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "'"

                SQL = "select * from View_Delay_Rpt where Location ='Blocked 2062' and Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' "

            End If

 
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(T01) Then
                X = X + 1
                _FirstChr = X

                worksheet1.Cells(X, 1) = "Blocked 2062"
                worksheet1.Range("A" & X & ":B" & X).MergeCells = True
                worksheet1.Range("A" & X & ":B" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                '  worksheet1.Range("A" & X, "A" & X).Interior.Color = RGB(141, 180, 227)
                worksheet1.Rows(X).Font.size = 10
                worksheet1.Rows(X).Font.name = "Tahoma"
                worksheet1.Rows(X).Font.Bold = True

                _Chart = 97
                For Y = 1 To 19

                    worksheet1.Range(ChrW(_Chart) & X - 1, ChrW(_Chart) & X - 1).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X - 1, ChrW(_Chart) & X - 1).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X - 1, ChrW(_Chart) & X - 1).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X - 1, ChrW(_Chart) & X - 1).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                    _Chart = _Chart + 1
                Next
                X = X + 1
                _FirstChr = X
                i = 0
                For Each DTRow3 As DataRow In T01.Tables(0).Rows
                    Dim _Tollaranz As Double
                    worksheet1.Rows(X).Font.size = 8
                    worksheet1.Rows(X).Font.name = "Tahoma"

                    _Tollaranz = 0
                    _Tollaranz = T01.Tables(0).Rows(i)("M01SO_Qty") * (T01.Tables(0).Rows(i)("M01Cus_Tol_Min") / 100)
                    '   MsgBox(CDbl(T01.Tables(0).Rows(i)("Order_Qty_mtr")))

                    worksheet1.Cells(X, 1) = T01.Tables(0).Rows(i)("Sales_Order")
                    worksheet1.Cells(X, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 2) = T01.Tables(0).Rows(i)("Line_Item")
                    worksheet1.Cells(X, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 3) = T01.Tables(0).Rows(i)("M01Cuatomer_Name")
                    worksheet1.Cells(X, 4) = T01.Tables(0).Rows(i)("Material")
                    worksheet1.Cells(X, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 5) = T01.Tables(0).Rows(i)("Material_Dis")
                    worksheet1.Cells(X, 6) = T01.Tables(0).Rows(i)("No_Dealy_Day")
                    worksheet1.Cells(X, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 7) = T01.Tables(0).Rows(i)("Delivery_Date")
                    worksheet1.Cells(X, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 8) = T01.Tables(0).Rows(i)("Lst_Confirn_Date")
                    worksheet1.Cells(X, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 9) = T01.Tables(0).Rows(i)("No_Day_Same_Opp")
                    worksheet1.Cells(X, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 10) = T01.Tables(0).Rows(i)("Product_Order")
                    worksheet1.Cells(X, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 11) = T01.Tables(0).Rows(i)("Order_Qty_Kg")
                    ' worksheet1.Cells(X, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(X, 11)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(X, 12) = T01.Tables(0).Rows(i)("Order_Qty_mtr")
                    ' worksheet1.Cells(X, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(X, 11)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(X, 13) = T01.Tables(0).Rows(i)("Status")
                    worksheet1.Cells(X, 14) = T01.Tables(0).Rows(i)("Pln_Comment")
                    worksheet1.Cells(X, 15) = T01.Tables(0).Rows(i)("NC_Comment")
                    worksheet1.Cells(X, 16) = T01.Tables(0).Rows(i)("Delay_reason")
                    worksheet1.Cells(X, 17) = T01.Tables(0).Rows(i)("Order_Type")
                    worksheet1.Cells(X, 17).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 18) = T01.Tables(0).Rows(i)("Brandix_OTD_Week")
                    worksheet1.Cells(X, 19) = T01.Tables(0).Rows(i)("Location")
                    worksheet1.Cells(X, 19).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    _Chart = 97
                    For Y = 1 To 19


                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                        _Chart = _Chart + 1
                    Next

                    X = X + 1

                    i = i + 1
                Next

                _2065 = X
                worksheet1.Cells(X, 3) = "Sum of Blocked 2062 "
                worksheet1.Range("C" & X, "C" & X).Interior.Color = RGB(141, 180, 227)
                worksheet1.Rows(X).Font.size = 10
                worksheet1.Rows(X).Font.name = "Tahoma"
                worksheet1.Rows(X).Font.Bold = True

                worksheet1.Range("K" & (X)).Formula = "=SUM(K" & _FirstChr & ":K" & (X - 1) & ")"
                worksheet1.Cells(X, 11).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet1.Cells(X, 11)
                range1.NumberFormat = "0.00"
                worksheet1.Range("K" & X, "K" & X).Interior.Color = RGB(141, 180, 227)

                worksheet1.Range("l" & (X)).Formula = "=SUM(L" & _FirstChr & ":L" & (X - 1) & ")"
                worksheet1.Cells(X, 12).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet1.Cells(X, 12)
                range1.NumberFormat = "0.00"
                worksheet1.Range("L" & X, "L" & X).Interior.Color = RGB(141, 180, 227)

                _Chart = 97
                For Y = 1 To 19

                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                    _Chart = _Chart + 1
                Next
            End If
            '====================================================================
            '2070
            If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                 "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                      "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                      "WHERE Z.Location='Blocked 2070' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') AND S.Department IN ('" & _Department & "') AND S.Merchant IN ('" & _Merchant & "') "
            ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                      "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                      "WHERE Z.Location='Blocked 2070' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') AND S.Department IN ('" & _Department & "')  "
            ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                        "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                        "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                        "WHERE Z.Location='Blocked 2070' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') AND S.Merchant IN ('" & _Merchant & "')  "
            ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                        "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                        "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                        "WHERE Z.Location='Blocked 2070' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND S.Department IN ('" & _Department & "') AND S.Merchant IN ('" & _Merchant & "')  "
            ElseIf Trim(txtCustomer.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                                "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                                "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                                "WHERE Z.Location='Blocked 2070' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND M01Cuatomer_Name IN ('" & _Customer & "') "
            ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                         "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                         "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                         "WHERE Z.Location='Blocked 2070' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "'  AND S.Department IN ('" & _Department & "')  "
            ElseIf Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                      "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                      "WHERE Z.Location='Blocked 2070' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "'  AND S.Merchant IN ('" & _Merchant & "') "

            Else
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,M01Cuatomer_Name " & _
                     " FROM ZPP_DEL z INNER JOIN  M01Sales_Order_SAP ON Z.Sales_Order=M01Sales_Order AND M01Line_Item=Z.Line_Item " & _
                    "left JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                    "WHERE Z.Location='Blocked 2070' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "'"

                '  SQL = "select * from View_Delay_Rpt where Location ='Blocked 2062' and Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' "

            End If

            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(T01) Then
                X = X + 1
                _FirstChr = X

                worksheet1.Cells(X, 1) = "Blocked 2070"
                worksheet1.Range("A" & X & ":B" & X).MergeCells = True
                worksheet1.Range("A" & X & ":B" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                '  worksheet1.Range("A" & X, "A" & X).Interior.Color = RGB(141, 180, 227)
                worksheet1.Rows(X).Font.size = 10
                worksheet1.Rows(X).Font.name = "Tahoma"
                worksheet1.Rows(X).Font.Bold = True

                _Chart = 97
                For Y = 1 To 19

                    worksheet1.Range(ChrW(_Chart) & X - 1, ChrW(_Chart) & X - 1).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X - 1, ChrW(_Chart) & X - 1).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X - 1, ChrW(_Chart) & X - 1).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X - 1, ChrW(_Chart) & X - 1).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                    _Chart = _Chart + 1
                Next
                X = X + 1
                _FirstChr = X
                i = 0
                For Each DTRow3 As DataRow In T01.Tables(0).Rows
                    Dim _Tollaranz As Double
                    worksheet1.Rows(X).Font.size = 8
                    worksheet1.Rows(X).Font.name = "Tahoma"

                    _Tollaranz = 0
                    _Tollaranz = T01.Tables(0).Rows(i)("M01SO_Qty") * (T01.Tables(0).Rows(i)("M01Cus_Tol_Min") / 100)
                    '   MsgBox(CDbl(T01.Tables(0).Rows(i)("Order_Qty_mtr")))

                    worksheet1.Cells(X, 1) = T01.Tables(0).Rows(i)("Sales_Order")
                    worksheet1.Cells(X, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 2) = T01.Tables(0).Rows(i)("Line_Item")
                    worksheet1.Cells(X, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 3) = T01.Tables(0).Rows(i)("M01Cuatomer_Name")
                    worksheet1.Cells(X, 4) = T01.Tables(0).Rows(i)("Material")
                    worksheet1.Cells(X, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 5) = T01.Tables(0).Rows(i)("Material_Dis")
                    worksheet1.Cells(X, 6) = T01.Tables(0).Rows(i)("No_Dealy_Day")
                    worksheet1.Cells(X, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 7) = T01.Tables(0).Rows(i)("Delivery_Date")
                    worksheet1.Cells(X, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 8) = T01.Tables(0).Rows(i)("Lst_Confirn_Date")
                    worksheet1.Cells(X, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 9) = T01.Tables(0).Rows(i)("No_Day_Same_Opp")
                    worksheet1.Cells(X, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 10) = T01.Tables(0).Rows(i)("Product_Order")
                    worksheet1.Cells(X, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 11) = T01.Tables(0).Rows(i)("Order_Qty_Kg")
                    ' worksheet1.Cells(X, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(X, 11)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(X, 12) = T01.Tables(0).Rows(i)("Order_Qty_mtr")
                    ' worksheet1.Cells(X, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(X, 11)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(X, 13) = T01.Tables(0).Rows(i)("Status")
                    worksheet1.Cells(X, 14) = T01.Tables(0).Rows(i)("Pln_Comment")
                    worksheet1.Cells(X, 15) = T01.Tables(0).Rows(i)("NC_Comment")
                    worksheet1.Cells(X, 16) = T01.Tables(0).Rows(i)("Delay_reason")
                    worksheet1.Cells(X, 17) = T01.Tables(0).Rows(i)("Order_Type")
                    worksheet1.Cells(X, 17).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 18) = T01.Tables(0).Rows(i)("Brandix_OTD_Week")
                    worksheet1.Cells(X, 19) = T01.Tables(0).Rows(i)("Location")
                    worksheet1.Cells(X, 19).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    _Chart = 97
                    For Y = 1 To 19


                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                        _Chart = _Chart + 1
                    Next

                    X = X + 1

                    i = i + 1
                Next

                _2065 = X
                worksheet1.Cells(X, 3) = "Sum of Blocked 2070 "
                worksheet1.Range("C" & X, "C" & X).Interior.Color = RGB(141, 180, 227)
                worksheet1.Rows(X).Font.size = 10
                worksheet1.Rows(X).Font.name = "Tahoma"
                worksheet1.Rows(X).Font.Bold = True

                worksheet1.Range("K" & (X)).Formula = "=SUM(K" & _FirstChr & ":K" & (X - 1) & ")"
                worksheet1.Cells(X, 11).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet1.Cells(X, 11)
                range1.NumberFormat = "0.00"
                worksheet1.Range("K" & X, "K" & X).Interior.Color = RGB(141, 180, 227)

                worksheet1.Range("l" & (X)).Formula = "=SUM(L" & _FirstChr & ":L" & (X - 1) & ")"
                worksheet1.Cells(X, 12).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet1.Cells(X, 12)
                range1.NumberFormat = "0.00"
                worksheet1.Range("L" & X, "L" & X).Interior.Color = RGB(141, 180, 227)

                _Chart = 97
                For Y = 1 To 19

                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                    _Chart = _Chart + 1
                Next
            End If
            '------------------------------------------------------------------------------------
            X = X + 2

            worksheet1.Cells(X, 3) = "Total Quantity"
            worksheet1.Range("C" & X, "C" & X).Interior.Color = RGB(141, 180, 227)
            worksheet1.Rows(X).Font.size = 10
            worksheet1.Rows(X).Font.name = "Tahoma"
            worksheet1.Rows(X).Font.Bold = True
            If _Aw_Dyeing = 0 Then
                _Aw_Dyeing = 5
            End If
            If _Aw_Exam = 0 Then
                _Aw_Exam = 5
            End If
            If _Nc1 = 0 Then
                _Nc1 = 5
            End If

            If _NCApp = 0 Then
                _NCApp = 5
            End If

            If _2062 = 0 Then
                _2062 = 5
            End If
            If _2070 = 0 Then
                _2070 = 5
            End If
            If _2065 = 0 Then
                _2065 = 5
            End If

            If _Aw_Finishing = 0 Then
                _Aw_Finishing = 5
            End If
            If _Aw_Grage = 0 Then
                _Aw_Grage = 5
            End If
            If _Awprep = 0 Then
                _Awprep = 5
            End If

            If _NC_FINISHING = 0 Then
                _NC_FINISHING = 5
            End If

            If _NC_KNT = 0 Then
                _NC_KNT = 5
            End If
            If _AWCONTAMI = 0 Then
                _AWCONTAMI = 5
            End If
            If _TOTALDYE = 0 Then
                _TOTALDYE = 5
            End If
            If _AWPrint = 0 Then
                _AWPrint = 5
            End If

            worksheet1.Range("K" & (X)).Formula = "=(K" & _Aw_Dyeing & "+K" & _Aw_Exam & "+K" & _Nc1 & "+K" & _NCApp & "+k" & _2062 & "+K" & _2065 & "+K" & _2070 & "+K" & _Aw_Finishing & "+K" & _Aw_Grage & "+K" & _Awprep & "+K" & _NC_FINISHING & "+K" & _NC_KNT & "+K" & _AWCONTAMI & "+K" & _TOTALDYE & "+k" & _AWPrint & ")"
            worksheet1.Cells(X, 11).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 11)
            range1.NumberFormat = "0.00"
            worksheet1.Range("K" & X, "K" & X).Interior.Color = RGB(141, 180, 227)

            worksheet1.Range("l" & (X)).Formula = "=(L" & _Aw_Dyeing & "+L" & _Aw_Exam & "+L" & _Nc1 & "+L" & _NCApp & "+L" & _2062 & "+L" & _2065 & "+L" & _2070 & "+L" & _Aw_Finishing & "+L" & _Aw_Grage & "+L" & _Awprep & "+L" & _NC_FINISHING & "+L" & _NC_KNT & "+L" & _AWCONTAMI & "+L" & _TOTALDYE & "+L" & _AWPrint & ")"
            worksheet1.Cells(X, 12).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 12)
            range1.NumberFormat = "0.00"
            worksheet1.Range("L" & X, "L" & X).Interior.Color = RGB(141, 180, 227)

            _Chart = 97
            For Y = 1 To 19
                worksheet1.Range(ChrW(_Chart) & X - 2, ChrW(_Chart) & X - 2).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet1.Range(ChrW(_Chart) & X - 2, ChrW(_Chart) & X - 2).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                worksheet1.Range(ChrW(_Chart) & X - 2, ChrW(_Chart) & X - 2).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                worksheet1.Range(ChrW(_Chart) & X - 2, ChrW(_Chart) & X - 2).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous


                worksheet1.Range(ChrW(_Chart) & X - 1, ChrW(_Chart) & X - 1).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet1.Range(ChrW(_Chart) & X - 1, ChrW(_Chart) & X - 1).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                worksheet1.Range(ChrW(_Chart) & X - 1, ChrW(_Chart) & X - 1).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                worksheet1.Range(ChrW(_Chart) & X - 1, ChrW(_Chart) & X - 1).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                _Chart = _Chart + 1
            Next

            X = X + 5
            worksheet1.Cells(X, 1) = "System Genarated Report by PPS " & Now
            worksheet1.Rows(X).Font.size = 10
            worksheet1.Rows(X).Font.name = "Calibri"
            worksheet1.Rows(X).Font.bold = True
            worksheet1.Rows(7).Select()
            worksheet1.Application.ActiveWindow.FreezePanes = True
            '=======================================================================================================>>>
            '==================================================================================>>
            '================================================>>
            '==================================================================================>>
            '========================================================================================================>>
            'SHEET 02

            workbooks.Application.Sheets.Add()
            Dim worksheet117 As _Worksheet = CType(sheets1.Item(1), _Worksheet)
            worksheet117.Name = "Order Delivery Info"
            worksheet117.Columns("A").ColumnWidth = 10
            worksheet117.Columns("B").ColumnWidth = 10
            worksheet117.Columns("C").ColumnWidth = 10
            worksheet117.Columns("D").ColumnWidth = 10
            worksheet117.Columns("E").ColumnWidth = 28
            worksheet117.Columns("F").ColumnWidth = 16
            worksheet117.Columns("G").ColumnWidth = 10
            worksheet117.Columns("H").ColumnWidth = 19

            'worksheet117.Columns("H").ColumnWidth = 18

            worksheet117.Cells(2, 1) = "Advance Sales Order Delivery Info"
            worksheet117.Rows(2).Font.size = 10
            worksheet117.Rows(2).Font.BOLD = True

            worksheet117.Range("A2:H2").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet117.Range(worksheet117.Cells(1, 1), worksheet117.Cells(1, 8)).Merge()
            worksheet117.Range(worksheet117.Cells(1, 1), worksheet117.Cells(1, 8)).HorizontalAlignment = XlHAlign.xlHAlignCenter


            'worksheet117.Cells(1, 10) = "Advance Sales Order Delivery Info"

            worksheet117.Range("J2:P2").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet117.Range(worksheet117.Cells(1, 10), worksheet117.Cells(1, 16)).Merge()
            worksheet117.Range(worksheet117.Cells(1, 10), worksheet117.Cells(1, 16)).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet117.Rows(2).Font.size = 12
            worksheet117.Rows(2).Font.BOLD = True
            worksheet1.Rows(2).Font.name = "Tahoma"
            worksheet1.Rows(2).rowheight = 24.25

            worksheet117.Cells(4, 1) = "Sales Order"
            worksheet117.Cells(4, 2) = "Line Item"
            worksheet117.Cells(4, 3) = "Material"
            worksheet117.Cells(4, 4) = "Material Description"
            worksheet117.Cells(4, 5) = "Delivery Date"
            worksheet117.Cells(4, 6) = "To be Pln Qty(M)"
            worksheet117.Cells(4, 7) = "Sales Order Qty"



            worksheet117.Cells(4, 9) = "Sales Order"
            worksheet117.Cells(4, 10) = "Line Item"
            worksheet117.Cells(4, 11) = "Material"
            worksheet117.Cells(4, 12) = "Material Description"
            worksheet117.Cells(4, 13) = "Delivery Date"
            worksheet117.Cells(4, 14) = "Sales Order Qty"
            worksheet117.Cells(4, 15) = "Delived Qty"

            worksheet117.Columns("A").ColumnWidth = 15
            worksheet117.Columns("B").ColumnWidth = 12
            worksheet117.Columns("C").ColumnWidth = 18
            worksheet117.Columns("D").ColumnWidth = 31
            worksheet117.Columns("E").ColumnWidth = 15
            worksheet117.Columns("F").ColumnWidth = 10
            worksheet117.Columns("M").ColumnWidth = 20
            worksheet117.Columns("G").ColumnWidth = 18
            worksheet117.Columns("H").ColumnWidth = 6
            worksheet117.Columns("N").ColumnWidth = 18
            worksheet117.Columns("J").ColumnWidth = 20

            worksheet117.Columns("I").ColumnWidth = 15
            worksheet117.Columns("J").ColumnWidth = 12
            worksheet117.Columns("K").ColumnWidth = 18
            worksheet117.Columns("L").ColumnWidth = 31
            worksheet117.Columns("M").ColumnWidth = 15
            worksheet117.Columns("N").ColumnWidth = 10

            worksheet117.Rows(4).Font.size = 10
            worksheet117.Rows(4).Font.BOLD = True
            worksheet1.Rows(4).Font.name = "Tahoma"
            worksheet1.Rows(4).rowheight = 24.25
            X = 4
            _Chart = 97
            For i = 1 To 7
                ' MsgBox(ChrW(_Chart))
                worksheet117.Range(UCase(ChrW(_Chart)) & X, UCase(ChrW(_Chart)) & X).Interior.Color = RGB(141, 180, 227)
                worksheet117.Cells(X, i).WrapText = True
                worksheet117.Range(UCase(ChrW(_Chart)) & X, UCase(ChrW(_Chart)) & X).MergeCells = True
                worksheet117.Range(UCase(ChrW(_Chart)) & X, UCase(ChrW(_Chart)) & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                worksheet117.Cells(X, i).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet117.Range(UCase(ChrW(_Chart)) & X, UCase(ChrW(_Chart)) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet117.Range(UCase(ChrW(_Chart)) & X, UCase(ChrW(_Chart)) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                worksheet117.Range(UCase(ChrW(_Chart)) & X, UCase(ChrW(_Chart)) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                worksheet117.Range(UCase(ChrW(_Chart)) & X, UCase(ChrW(_Chart)) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                _Chart = _Chart + 1
            Next
            _Chart = _Chart + 1
            '  MsgBox(UCase(ChrW(_Chart)))
            Y = 9
            For i = 1 To 7
                ' MsgBox(ChrW(_Chart))
                worksheet117.Range(UCase(ChrW(_Chart)) & X, UCase(ChrW(_Chart)) & X).Interior.Color = RGB(141, 180, 227)
                worksheet117.Cells(X, Y).WrapText = True
                worksheet117.Range(UCase(ChrW(_Chart)) & X, UCase(ChrW(_Chart)) & X).MergeCells = True
                worksheet117.Range(UCase(ChrW(_Chart)) & X, UCase(ChrW(_Chart)) & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                worksheet117.Cells(X, Y).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet117.Range(UCase(ChrW(_Chart)) & X, UCase(ChrW(_Chart)) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet117.Range(UCase(ChrW(_Chart)) & X, UCase(ChrW(_Chart)) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                worksheet117.Range(UCase(ChrW(_Chart)) & X, UCase(ChrW(_Chart)) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                worksheet117.Range(UCase(ChrW(_Chart)) & X, UCase(ChrW(_Chart)) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                Y = Y + 1
                _Chart = _Chart + 1
            Next

            Dim _Second As Integer

            If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,s.Customer " & _
                      "INNER JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                      "WHERE Z.WIP_Loc='To Be planned' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND S.Customer IN ('" & _Customer & "') AND S.Department IN ('" & _Department & "') AND S.Merchant IN ('" & _Merchant & "') "
            ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,s.Customer " & _
                      "INNER JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                      "WHERE Z.WIP_Loc='To Be planned' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND S.Customer IN ('" & _Customer & "') AND S.Department IN ('" & _Department & "')  "
            ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                        "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,s.Customer " & _
                        "INNER JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                        "WHERE Z.WIP_Loc='To Be planned' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND S.Customer IN ('" & _Customer & "') AND S.Merchant IN ('" & _Merchant & "')  "
            ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                        "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,s.Customer " & _
                        "INNER JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                        "WHERE Z.WIP_Loc='To Be planned' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND S.Department IN ('" & _Department & "') AND S.Merchant IN ('" & _Merchant & "')  "
            ElseIf Trim(txtCustomer.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                                "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,s.Customer " & _
                                "INNER JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                                "WHERE Z.WIP_Loc='To Be planned' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "' AND S.Customer IN ('" & _Customer & "') "
            ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                         "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,s.Customer " & _
                         "INNER JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                         "WHERE Z.WIP_Loc='To Be planned' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "'  AND S.Department IN ('" & _Department & "')  "
            ElseIf Trim(txtMerchant.Text) <> "" Then
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,s.Customer " & _
                      "INNER JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                      "WHERE Z.WIP_Loc='To Be planned' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "'  AND S.Merchant IN ('" & _Merchant & "') "

            Else
                SQL = "SELECT z.Sales_Order,z.Line_Item,z.Material,z.Material_Dis,z.No_Dealy_Day,z.Delivery_Date,z.Lst_Confirn_Date,z.No_Day_Same_Opp,z.WIP_Loc,z.Product_Order,z.Order_Qty_mtr, " & _
                      "z.Order_Qty_Kg,z.Status,z.Pln_Comment,z.NC_Comment,z.Delay_reason,z.Order_Type,z.Brandix_OTD_Week,z.Location,M01Cus_Tol_Min,M01SO_Qty,s.Department,s.Merchant,s.Customer " & _
                     " FROM ZPP_DEL z INNER JOIN  M01Sales_Order_SAP ON Z.Sales_Order=M01Sales_Order AND M01Line_Item=Z.Line_Item " & _
                    "INNER JOIN OTD_SMS S ON S.Sales_Order=M01Sales_Order AND S.Line_Item=M01Line_Item and s.Del_Date=z.Delivery_Date " & _
                    "WHERE Z.WIP_Loc='To Be planned' AND Z.Delivery_Date BETWEEN '" & _FROMDATE & "' AND '" & _TODATE & "'"
            End If

            X = X + 1
            _Second = X
            _FirstChr = X
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            i = 0
            For Each DTRow3 As DataRow In T01.Tables(0).Rows
                Dim _Tollaranz As Double
                worksheet117.Rows(X).Font.size = 8
                worksheet117.Rows(X).Font.name = "Tahoma"

                _Tollaranz = 0
                _Tollaranz = T01.Tables(0).Rows(i)("M01SO_Qty") * (T01.Tables(0).Rows(i)("M01Cus_Tol_Min") / 100)
                '   MsgBox(CDbl(T01.Tables(0).Rows(i)("Order_Qty_mtr")))
                If _Tollaranz < CDbl(T01.Tables(0).Rows(i)("Order_Qty_mtr")) Then
                    worksheet117.Cells(X, 1) = T01.Tables(0).Rows(i)("Sales_Order")
                    worksheet117.Cells(X, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet117.Cells(X, 2) = T01.Tables(0).Rows(i)("Line_Item")
                    worksheet117.Cells(X, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet117.Cells(X, 3) = T01.Tables(0).Rows(i)("Material")
                    worksheet117.Cells(X, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet117.Cells(X, 4) = T01.Tables(0).Rows(i)("Material_Dis")

                    worksheet117.Cells(X, 5) = T01.Tables(0).Rows(i)("Delivery_Date")
                    worksheet117.Cells(X, 6) = T01.Tables(0).Rows(i)("Order_Qty_mtr")
                    ' worksheet1.Cells(X, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet117.Cells(X, 6)
                    range1.NumberFormat = "0.00"
                    worksheet117.Cells(X, 7) = T01.Tables(0).Rows(i)("m01SO_QTY")
                    ' worksheet1.Cells(X, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet117.Cells(X, 7)
                    range1.NumberFormat = "0.00"

                    SQL = "select max(Metrrial) as Metrrial,max(Met_Des) as Met_Des,Del_Date,Sales_Order,Line_Item,sum(Del_Qty) as Del_Qty from OTD_Records where Metrrial='" & Trim(T01.Tables(0).Rows(i)("Material")) & "' and Del_Date> '" & Trim(T01.Tables(0).Rows(i)("Delivery_Date")) & "'  and Del_Qty>0 group by Sales_Order,Line_Item,Del_Date"
                    T03 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                    Dim z As Integer
                    z = 0
                    For Each DTRow4 As DataRow In T03.Tables(0).Rows
                        worksheet117.Cells(_Second, 9) = T03.Tables(0).Rows(z)("Sales_Order")
                        worksheet117.Cells(_Second, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        worksheet117.Cells(_Second, 10) = T03.Tables(0).Rows(z)("Line_Item")
                        worksheet117.Cells(_Second, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        worksheet117.Cells(_Second, 11) = T03.Tables(0).Rows(z)("Metrrial")
                        worksheet117.Cells(_Second, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        worksheet117.Cells(_Second, 12) = T03.Tables(0).Rows(z)("Met_Des")
                        worksheet117.Cells(_Second, 13) = T03.Tables(0).Rows(z)("Del_Date")
                        worksheet117.Cells(_Second, 13).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        worksheet117.Cells(_Second, 14) = T01.Tables(0).Rows(i)("m01SO_QTY")
                        ' worksheet1.Cells(X, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        range1 = worksheet117.Cells(_Second, 14)
                        range1.NumberFormat = "0.00"
                        worksheet117.Cells(_Second, 15) = T03.Tables(0).Rows(z)("Del_Qty")
                        ' worksheet1.Cells(X, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        range1 = worksheet117.Cells(_Second, 15)
                        range1.NumberFormat = "0.00"

                        _Chart = 105
                        For Y = 1 To 7


                            worksheet117.Range(ChrW(_Chart) & _Second, ChrW(_Chart) & _Second).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlDot
                            worksheet117.Range(ChrW(_Chart) & _Second, ChrW(_Chart) & _Second).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDot
                            worksheet117.Range(ChrW(_Chart) & _Second, ChrW(_Chart) & _Second).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlDot
                            worksheet117.Range(ChrW(_Chart) & _Second, ChrW(_Chart) & _Second).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDot

                            _Chart = _Chart + 1
                        Next
                        _Second = _Second + 1
                        z = z + 1
                    Next

                    _Chart = 97
                    For Y = 1 To 7


                        worksheet117.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlDot
                        worksheet117.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDot
                        worksheet117.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlDot
                        worksheet117.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDot

                        _Chart = _Chart + 1
                    Next

                    X = X + 1
                End If
                i = i + 1
            Next

            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            ' worksheet1.Cells(4, 5) = _Fail_Batch
            'worksheet1.Cells(4, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
            MsgBox("Report Genarated successfully", MsgBoxStyle.Information, "Technova ....")
            ' MsgBox(_Fail_Batch)
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
            End If
        End Try
    End Function
End Class