Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports CrystalDecisions.CrystalReports.Engine

Imports System.IO
Public Class frmItem_Issue_Uniq
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
    Dim _INEmp1 As String
    Dim _INEmp2 As String
    Dim _permissionLevel As String
    Dim _lastRow As Integer

    Function Load_Gride2()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTable_itemIssue_Uniq
        UltraGrid1.DataSource = c_dataCustomer1
        With UltraGrid1
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

    Function Load_Gride3()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer2 = CustomerDataClass.MakeDataTable_iNVOICE_UNIQ
        UltraGrid2.DataSource = c_dataCustomer2
        With UltraGrid2
            .DisplayLayout.Bands(0).Columns(0).Width = 110
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 170
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(3).Width = 70
            .DisplayLayout.Bands(0).Columns(4).Width = 70
            .DisplayLayout.Bands(0).Columns(5).Width = 70
            .DisplayLayout.Bands(0).Columns(6).Width = 70
            .DisplayLayout.Bands(0).Columns(7).Width = 110


            .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            '.DisplayLayout.Bands(0).Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            .DisplayLayout.Bands(0).Columns(0).CellActivation = Activation.NoEdit
            .DisplayLayout.Bands(0).Columns(1).CellActivation = Activation.NoEdit
            .DisplayLayout.Bands(0).Columns(2).CellActivation = Activation.NoEdit
            .DisplayLayout.Bands(0).Columns(3).CellActivation = Activation.NoEdit
            .DisplayLayout.Bands(0).Columns(4).CellActivation = Activation.NoEdit
            .DisplayLayout.Bands(0).Columns(5).CellActivation = Activation.NoEdit
            .DisplayLayout.Bands(0).Columns(6).CellActivation = Activation.NoEdit
            '.DisplayLayout.Bands(0).Columns(7).CellActivation = Activation.NoEdit


            '.DisplayLayout.Bands(0).Columns(0).CellActivation = Activation.NoEdit
            '  .DisplayLayout.Bands(0).Columns(6).Width = 90
            ' .DisplayLayout.Bands(0).Columns(7).Width = 90

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function

    Function Load_DEPARTMENT()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select M08Description as [##] from M08Department where M08Status='A' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboDepartment
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 222
                ' .Rows.Band.Columns(1).Width = 360

            End With

            With cboDepartment1
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 180
                ' .Rows.Band.Columns(1).Width = 360

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

    Function Search_Invoice()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Dim i As Integer
        Dim VALUE As Double
        Dim _sT As String
        Try
            Sql = "SELECT * FROM T08Sales_Header WHERE T08Invo_No='" & Trim(txtEnter1.Text) & "' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                txtDate1.Text = M01.Tables(0).Rows(0)("T08Date")
                cboJob1.Text = Trim(M01.Tables(0).Rows(0)("T08Job_No"))
                cboVno1.Text = Trim(M01.Tables(0).Rows(0)("T08V_No"))
                Call Search_Vehicle_No_1()
                txtMtr1.Text = Trim(M01.Tables(0).Rows(0)("T08St_Mtr"))
                cboService.Text = Trim(M01.Tables(0).Rows(0)("T08Service_on"))
                ToolStripMenuItem2.Enabled = True
                cmdDelete1.Enabled = True
            End If

            Sql = "select M10Name from T10Technicion_Comm inner join M10Employee on M10Code=T10Emp where T10INV_No='" & Trim(txtEnter1.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            For Each DTRow2 As DataRow In M01.Tables(0).Rows
                If i = 0 Then
                    cboEmp01.Text = Trim(M01.Tables(0).Rows(0)("M10Name"))
                Else
                    cboEmp02.Text = Trim(M01.Tables(0).Rows(0)("M10Name"))
                End If
                i = i + 1
            Next
            '================================================================================
            Sql = "select * from T09Sales_Flutter  where T09Inv_No='" & Trim(txtEnter1.Text) & "' and T09Department<>'-'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            For Each DTRow2 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer2.NewRow
                newRow("Department") = Trim(M01.Tables(0).Rows(i)("T09Department"))
                newRow("#Part No") = Trim(M01.Tables(0).Rows(i)("T09Item_Code"))
                newRow("Item Name") = Trim(M01.Tables(0).Rows(i)("T09Item_Name"))
                VALUE = M01.Tables(0).Rows(i)("T09Retail")
                _sT = (VALUE.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _sT = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", VALUE))
                newRow("Retail Price") = _sT
                newRow("Qty") = M01.Tables(0).Rows(i)("T09Qty")
                newRow("Free Issue") = M01.Tables(0).Rows(i)("T09Free")
                newRow("Discount%") = M01.Tables(0).Rows(i)("T09Discount")
                VALUE = (CDbl(M01.Tables(0).Rows(i)("T09Retail")) * CDbl(M01.Tables(0).Rows(i)("T09Qty")))
                _sT = (VALUE.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _sT = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", VALUE))
                newRow("Total") = _sT
                ' _TOTAL = VALUE + _TOTAL
                c_dataCustomer2.Rows.Add(newRow)
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

    Private Sub frmItem_Issue_Uniq_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        frmView_Item_Uniq.Close()
        frmPayMain_uniq.Close()
        frmView_Invoice_Job.Close()
    End Sub

    Private Sub frmItem_Issue_Uniq_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_Gride2()
        Call Load_Gride3()
        txtEntry.ReadOnly = True
        txtEntry.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtCount.ReadOnly = True
        txtCount.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtDiscount.ReadOnly = True
        txtDiscount.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtRate.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtRate.ReadOnly = True
        txtQty.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtTotal.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtTotal.ReadOnly = True
        txtTp.ReadOnly = True
        txtAddress.ReadOnly = True
        txtCustomer.ReadOnly = True
        txtDate.Text = Today
        txtMtr.ReadOnly = True
        txtMtr.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtFree.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        Call Load_EntryNo()
        Call Load_DEPARTMENT()
        Call Load_VNO()
        Call Load_JobNo()
        Call Load_Employee()
        ' Call Load_Item()
        Call Load_NEXT_MILLAGE()
        txtDiscount.Text = "0"

        txtEnter1.ReadOnly = True
        txtEnter1.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtNet1.ReadOnly = True
        txtNet1.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
        txtNext.ReadOnly = True
        txtDate1.Text = Today
        txtMtr1.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        Call Load_EntryNo1()
        txtCname1.ReadOnly = True
        txtAddress1.ReadOnly = True
        txtTP1.ReadOnly = True
        Call Load_Employee_1()
        txtNext.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtRate1.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtCount1.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        UltraGrid2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
    End Sub

    Function Load_Searvice()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select M21Service as [##] from M21Job_Services where M21Department='" & Trim(cboDepartment1.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboService1
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 535

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
    Function Load_Employee()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select M10Name as [##] from M10Employee where M10Status='A'  and M10Designation not in ('Cashier','Manager','Accountant')"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboEmp
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 222

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


    Function Load_NEXT_MILLAGE()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select M11Name as [##] from M11Common where M11Status='SRV'  "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboService
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 107

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


    Function Search_Emp() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select * from M10Employee where M10Status='A' and M10Name='" & Trim(cboEmp.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                Search_Emp = True
                _Emp = Trim(M01.Tables(0).Rows(0)("M10Code"))
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


    Function Search_Emp_1() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            _INEmp1 = ""
            Sql = "select * from M10Employee where M10Status='A' and M10Name='" & Trim(cboEmp01.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                Search_Emp_1 = True
                _INEmp1 = Trim(M01.Tables(0).Rows(0)("M10Code"))
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

    Function Search_Emp_2() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            _INEmp2 = ""
            Sql = "select * from M10Employee where M10Status='A' and M10Name='" & Trim(cboEmp02.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                Search_Emp_2 = True
                _INEmp2 = Trim(M01.Tables(0).Rows(0)("M10Code"))
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

    Function Load_VNO()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select T05Vehi_No as [##],M06Name as [Customer Name] from T05Job_Card inner join M06Customer_Master on M06Code=T05Cus_No where T05Status='A'  order by T05Id"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboV_no
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 159
                .Rows.Band.Columns(1).Width = 210

            End With

            With cboVno1
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

    Function Load_JobNo()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select T05Job_No as [##],T05Vehi_No as [Vehicle No],M06Name as [Customer Name] from T05Job_Card inner join M06Customer_Master on M06Code=T05Cus_No where T05Status='A'  order by T05Id"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboJob
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 222
                .Rows.Band.Columns(1).Width = 110
                .Rows.Band.Columns(2).Width = 260
            End With

            With cboJob1
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 222
                .Rows.Band.Columns(1).Width = 110
                .Rows.Band.Columns(2).Width = 260
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

    Function Load_Employee_1()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select M10Name as [##] from M10Employee where M10Status='A'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboEmp01
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 222


            End With
            With cboEmp02
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 222


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


    Function Load_Service_Discription()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select T09Item_Name as [## from T09Sales_Flutter inner join T08Sales_Header on T08Invo_No=T09Inv_No where T08Tr_Type='JOB_SALES' AND T09Item_Code='-' GROUP BY T09Item_Name"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboService1
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 535

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
            Sql = "select * from P01Parameter where  P01Code='ISU'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                If M01.Tables(0).Rows(0)("P01No") >= 1 And M01.Tables(0).Rows(0)("P01No") < 10 Then
                    txtEntry.Text = "ISU-00" & M01.Tables(0).Rows(0)("P01No")
                ElseIf M01.Tables(0).Rows(0)("P01No") >= 10 And M01.Tables(0).Rows(0)("P01No") < 100 Then
                    txtEntry.Text = "ISU-0" & M01.Tables(0).Rows(0)("P01No")
                Else
                    txtEntry.Text = "ISU-" & M01.Tables(0).Rows(0)("P01No")
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

    Function Load_EntryNo1()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select * from P01Parameter where  P01Code='INV'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                If M01.Tables(0).Rows(0)("P01No") >= 1 And M01.Tables(0).Rows(0)("P01No") < 10 Then
                    txtEnter1.Text = "INV-00" & M01.Tables(0).Rows(0)("P01No")
                ElseIf M01.Tables(0).Rows(0)("P01No") >= 10 And M01.Tables(0).Rows(0)("P01No") < 100 Then
                    txtEnter1.Text = "INV-0" & M01.Tables(0).Rows(0)("P01No")
                Else
                    txtEnter1.Text = "INV-" & M01.Tables(0).Rows(0)("P01No")
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

    Private Sub cboJob_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboJob.AfterCloseUp
        Call Search_Jobno()
    End Sub

    Private Sub cboJob_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboJob.KeyUp
        If e.KeyCode = 13 Then
            cboEmp.ToggleDropdown()
        End If
    End Sub

    Private Sub cboEmp_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboEmp.KeyUp
        If e.KeyCode = 13 Then
            cboDepartment.ToggleDropdown()
        End If
    End Sub

    Function Search_Vehicle_No() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Search_Vehicle_No = False
            '---------------------------------------------SEARCH BY JOB NO
            Sql = "SELECT T05Job_No,T05Mtr,T05Department FROM T05Job_Card WHERE T05Vehi_No='" & Trim(cboV_no.Text) & "' AND T05Status='A' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                cboJob.Text = Trim(M01.Tables(0).Rows(0)("T05Job_No"))
                txtMtr.Text = Trim(M01.Tables(0).Rows(0)("T05Mtr"))
                cboDepartment.Text = Trim(M01.Tables(0).Rows(0)("T05Department"))
                '  Call Search_Jobno()
            End If
            '====================================================================
            Sql = "select * from M07Vehicle_Master inner join M06Customer_Master on M06Code=M07Cus_Code where M07Status='A'  and M07V_No='" & Trim(cboV_no.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                Search_Vehicle_No = True
                _CusNo = Trim(M01.Tables(0).Rows(0)("M06Code"))
                'cboBrand.Text = Trim(M01.Tables(0).Rows(0)("M07Brand_Name"))
                'cbov_Type.Text = Trim(M01.Tables(0).Rows(0)("M07Type"))
                txtTp.Text = Trim(M01.Tables(0).Rows(0)("M06Mobile_No"))
                txtCustomer.Text = Trim(M01.Tables(0).Rows(0)("M06Name"))
                txtAddress.Text = Trim(M01.Tables(0).Rows(0)("M06Address"))
                ' cboCus_Type.Text = Trim(M01.Tables(0).Rows(0)("M06Cus_Type"))
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

    Function Search_Vehicle_No_1() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Search_Vehicle_No_1 = False
            Sql = "SELECT T05Job_No,T05Mtr FROM T05Job_Card WHERE T05Vehi_No='" & Trim(cboVno1.Text) & "' AND T05Status='A' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                cboJob1.Text = Trim(M01.Tables(0).Rows(0)("T05Job_No"))
                txtMtr1.Text = Trim(M01.Tables(0).Rows(0)("T05Mtr"))
                ' Call Search_Jobno_1()
            End If
            '=====================================================================
            Sql = "select * from M07Vehicle_Master inner join M06Customer_Master on M06Code=M07Cus_Code where M07Status='A'  and M07V_No='" & Trim(cboVno1.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                Search_Vehicle_No_1 = True
                _CusNo = Trim(M01.Tables(0).Rows(0)("M06Code"))
                'cboBrand.Text = Trim(M01.Tables(0).Rows(0)("M07Brand_Name"))
                'cbov_Type.Text = Trim(M01.Tables(0).Rows(0)("M07Type"))
                txtTP1.Text = Trim(M01.Tables(0).Rows(0)("M06Mobile_No"))
                txtCname1.Text = Trim(M01.Tables(0).Rows(0)("M06Name"))
                txtAddress1.Text = Trim(M01.Tables(0).Rows(0)("M06Address"))
                ' cboCus_Type.Text = Trim(M01.Tables(0).Rows(0)("M06Cus_Type"))
            End If
            Call lOAD_DATA_GRIDE()
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

    Function Search_Jobno() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Search_Jobno = False
            Sql = "select * from T05Job_Card where T05Status ='A'  and T05Job_No='" & Trim(cboJob.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                Search_Jobno = True
              
                cboV_no.Text = Trim(M01.Tables(0).Rows(0)("T05Vehi_No"))
                txtMtr.Text = Trim(M01.Tables(0).Rows(0)("T05Mtr"))
                cboDepartment.Text = Trim(M01.Tables(0).Rows(0)("T05Department"))
                Call Search_Vehicle_No()
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

    Function Search_Jobno_1() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Search_Jobno_1 = False
            _CusNo = ""
            Sql = "select * from T05Job_Card where T05Status ='A'  and T05Job_No='" & Trim(cboJob1.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                Search_Jobno_1 = True
                _CusNo = Trim(M01.Tables(0).Rows(0)("T05Cus_No"))
                cboVno1.Text = Trim(M01.Tables(0).Rows(0)("T05Vehi_No"))
                txtMtr1.Text = Trim(M01.Tables(0).Rows(0)("T05Mtr"))
                Call Search_Vehicle_No_1()
            End If

            Call Load_Gride3()
            Call lOAD_DATA_GRIDE()
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

    Function lOAD_DATA_GRIDE()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Dim Value As Double
        Dim I As Integer
        Dim _st As String
        Dim _TOTAL As Double
        Try

            Sql = "SELECT * FROM T07Item_Issue_Fluter INNER JOIN  T06Item_Issue_Header ON T06Ref_No=T07Ref_No WHERE T06Job_No='" & Trim(cboJob1.Text) & "' AND T06Status IN ('ISSUE','CLOSE')"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            I = 0
            For Each DTRow2 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer2.NewRow
                newRow("Department") = "-"
                newRow("#Part No") = Trim(M01.Tables(0).Rows(I)("T07Part_No"))
                newRow("Item Name") = Trim(M01.Tables(0).Rows(I)("T07Item_Name"))
                Value = M01.Tables(0).Rows(I)("T07Rate")
                _st = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _st = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Retail Price") = _st
                Value = M01.Tables(0).Rows(I)("T07Qty")
                _st = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _st = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Qty") = _st
                newRow("Free Issue") = M01.Tables(0).Rows(I)("T07Free")
                Value = M01.Tables(0).Rows(I)("T07Discount")
                _st = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _st = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Discount%") = _st
                Value = (CDbl(M01.Tables(0).Rows(I)("T07Rate")) * CDbl(M01.Tables(0).Rows(I)("T07Qty"))) - (CDbl(M01.Tables(0).Rows(I)("T07Rate")) * CDbl(M01.Tables(0).Rows(I)("T07Qty"))) * CDbl(M01.Tables(0).Rows(I)("T07Discount")) / 100
                _st = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _st = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Total") = _st
                _TOTAL = Value + _TOTAL
                c_dataCustomer2.Rows.Add(newRow)
                I = I + 1
            Next
            txtNet1.Text = (_TOTAL.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            txtNet1.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", _TOTAL))
            txtCount1.Text = UltraGrid2.Rows.Count

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

    Private Sub cboDepartment_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboDepartment.KeyUp
        If e.KeyCode = 13 Then
            cboPartno.ToggleDropdown()
        End If
    End Sub

    Function Calculation_total()
        On Error Resume Next
        Dim Value As Double
        If IsNumeric(txtRate.Text) And IsNumeric(txtQty.Text) And IsNumeric(txtDiscount.Text) Then
            Value = CDbl(txtRate.Text) * CDbl(txtQty.Text)
            Value = Value - (Value * CDbl(txtDiscount.Text) / 100)
            txtTotal.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            txtTotal.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
        End If
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
                cboPartno.Text = Trim(M01.Tables(0).Rows(0)("M05Item_Code"))
                txtItem.Text = Trim(M01.Tables(0).Rows(0)("tmpDescription"))
                Value = Trim(M01.Tables(0).Rows(0)("Retail"))
                txtRate.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtRate.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                txtRate.ReadOnly = True
            Else
                txtRate.Text = ""
                txtRate.ReadOnly = False
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

    Private Sub cboPartno_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboPartno.AfterCloseUp
        'Call Search_Item()
    End Sub

    Private Sub cboPartno_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboPartno.KeyUp
        If e.KeyCode = 13 Then
            If Trim(cboPartno.Text) <> "" Then
                If UltraGrid3.Visible = True Then
                    UltraGrid3.Focus()
                Else
                    txtQty.Focus()
                End If
            End If
        ElseIf e.KeyCode = Keys.Escape Then
            UltraGrid3.Visible = False
        End If
    End Sub

    Private Sub txtRate_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtRate.KeyUp
        Dim Value As Double
        If e.KeyCode = 13 Then
            If txtRate.Text <> "" Then
                If IsNumeric(txtRate.Text) Then
                    Value = txtRate.Text
                    txtRate.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    txtRate.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                End If
                txtQty.Focus()
            End If
        End If
    End Sub

    Private Sub txtRate_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRate.TextChanged
        Call Calculation_total()
    End Sub

    Private Sub txtQty_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtQty.KeyUp
        If e.KeyCode = 13 Then
            txtFree.Focus()
        End If
    End Sub

    Private Sub txtQty_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtQty.TextChanged
        Call Calculation_total()
    End Sub

    Private Sub txtFree_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtFree.KeyUp
        Try
            If e.KeyCode = 13 Then

                If Trim(cboPartno.Text) <> "" Then
                Else
                    MsgBox("Please enter the Part No", MsgBoxStyle.Information, "Information ........")
                    ' cboPartno.ToggleDropdown()
                    Exit Sub
                End If

                If Trim(txtItem.Text) <> "" Then
                Else
                    MsgBox("Please enter the Part Name", MsgBoxStyle.Information, "Information ........")
                    ' txtItem.Focus()
                    Exit Sub
                End If



                If txtRate.Text <> "" Then
                Else
                    txtRate.Text = "0"
                End If

                If IsNumeric(txtRate.Text) Then
                Else
                    MsgBox("Please enter the correct Rate", MsgBoxStyle.Information, "Information .........")
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
                newRow("#Part No") = Trim(cboPartno.Text)
                newRow("Item Name") = Trim(txtItem.Text)
                newRow("Retail Price") = txtRate.Text
                newRow("Qty") = txtQty.Text
                newRow("Free Issue") = txtFree.Text
                newRow("Discount%") = txtDiscount.Text
                newRow("Total") = txtTotal.Text
                c_dataCustomer1.Rows.Add(newRow)
                Me.cboPartno.Text = ""
                Me.txtItem.Text = ""
                Me.txtQty.Text = "0"
                Me.txtRate.Text = "00.00"
                Me.txtFree.Text = "0"
                Me.txtTotal.Text = "00.00"
                Me.txtDiscount.Text = "0"
                strItem_Code = ""

                cboPartno.ToggleDropdown()
                UltraGrid3.Visible = False
                txtCount.Text = UltraGrid1.Rows.Count
            End If
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Sub

    Private Sub cboV_no_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboV_no.AfterCloseUp
        Call Search_Vehicle_No()
    End Sub

    Private Sub cboV_no_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboV_no.KeyUp
        If e.KeyCode = 13 Then
            cboPartno.ToggleDropdown()
        End If
    End Sub

    Private Sub UltraGrid1_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles UltraGrid1.AfterRowsDeleted
        txtCount.Text = UltraGrid1.Rows.Count
    End Sub

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        If Search_Jobno() = True Then
        Else
            MsgBox("Please select the correct Job No", MsgBoxStyle.Information, "Information .........")
            cboJob.ToggleDropdown()
            Exit Sub
        End If

        If Search_Emp() = True Then
        Else
            MsgBox("Please select the correct Employee Name", MsgBoxStyle.Information, "Information ...........")
            cboEmp.ToggleDropdown()
            Exit Sub
        End If

        If Trim(cboDepartment.Text) <> "" Then
        Else
            MsgBox("Please select the Department", MsgBoxStyle.Information, "Information .......")
            cboDepartment.ToggleDropdown()
            Exit Sub
        End If

        If UltraGrid1.Rows.Count > 0 Then
        Else
            MsgBox("Please enter the Item detailes", MsgBoxStyle.Information, "Information .......")
            cboPartno.ToggleDropdown()
            Exit Sub
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
        Dim _GetDate As DateTime
        Dim _Get_Time As DateTime
        Dim A As String
        Dim B As New ReportDocument

        Dim M01 As DataSet
        Dim M02 As DataSet
        Dim I As Integer
        Try
            nvcFieldList1 = "select * from T06Item_Issue_Header where T06Ref_No='" & Trim(txtEntry.Text) & "'"
            M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(M01) Then
                'MsgBox("This Transaction No alrady exsist", MsgBoxStyle.Information, "Information .........")
                'connection.Close()
                'Exit Function

                nvcFieldList1 = "select * from T05Job_Card  where T05Job_No='" & Trim(cboJob.Text) & "' and T05Status='A' "
                M02 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                If isValidDataset(M02) Then
                    nvcFieldList1 = "UPDATE T06Item_Issue_Header SET T06Job_No='" & Trim(cboJob.Text) & "',T06Date='" & txtDate.Text & "',T06Emp='" & _Emp & "',T06Department='" & Trim(cboDepartment.Text) & "',T06V_no='" & Trim(cboV_no.Text) & "',T06Cus_No='" & _CusNo & "',T06Status='ISSUE' WHERE T06Ref_No='" & txtEntry.Text & "'"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                    nvcFieldList1 = "Insert Into tmpTransaction_Log(tmpStatus,tmpProcess,tmpTime,tmpUser,tmpCode)" & _
                                                  " values('ITEM ISSUE','EDIT', '" & Now & "','" & strDisname & "','" & Trim(txtEntry.Text) & "')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                    nvcFieldList1 = "DELETE FROM T07Item_Issue_Fluter WHERE T07Ref_No='" & txtEntry.Text & "'"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                    nvcFieldList1 = "DELETE FROM S01Stock_Balance WHERE S01Ref_No='" & txtEntry.Text & "' AND S01Tr_Type='ISSUE_ITEM'"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                    I = 0
                    For Each uRow As UltraGridRow In UltraGrid1.Rows
                        nvcFieldList1 = "SELECT * FROM M05Item_Master WHERE M05Item_Code ='" & Trim(UltraGrid1.Rows(I).Cells(0).Text) & "'"
                        M02 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                        If isValidDataset(M02) Then
                            nvcFieldList1 = "Insert Into S01Stock_Balance(S01Item_Code,S01Ref_No,S01Date,S01Time,S01Tr_Type,S01Qty,S01Status)" & _
                                                    " values('" & Trim(UltraGrid1.Rows(I).Cells(0).Text) & "','" & txtEntry.Text & "', '" & txtDate.Text & "','" & Now & "','ISSUE_ITEM','" & -(CDbl(UltraGrid1.Rows(I).Cells(4).Text) + CDbl(UltraGrid1.Rows(I).Cells(5).Text)) & "','A')"
                            ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                        End If
                        nvcFieldList1 = "Insert Into T07Item_Issue_Fluter(T07Ref_No,T07Part_No,T07Item_Name,T07Rate,T07Discount,T07Qty,T07Free,T07Status)" & _
                                              " values('" & txtEntry.Text & "','" & Trim(UltraGrid1.Rows(I).Cells(0).Text) & "','" & Trim(UltraGrid1.Rows(I).Cells(1).Text) & "', '" & Trim(UltraGrid1.Rows(I).Cells(2).Text) & "','" & Trim(UltraGrid1.Rows(I).Cells(3).Text) & "','" & Trim(UltraGrid1.Rows(I).Cells(4).Text) & "','" & Trim(UltraGrid1.Rows(I).Cells(5).Text) & "','A')"
                        ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                        I = I + 1
                    Next
                    MsgBox("Items issue successfully", MsgBoxStyle.Information, "Information ........")
                Else
                    MsgBox("This Job No alrady close", MsgBoxStyle.Information, "Information ........")
                    connection.ClearAllPools()
                    connection.Close()
                    Exit Function
                End If
            Else
                Call Load_EntryNo()

                nvcFieldList1 = "update P01Parameter set P01No=P01No+ " & 1 & " where P01Code='ISU' "
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                'SAVE ISSUE HEADER
                nvcFieldList1 = "Insert Into T06Item_Issue_Header(T06Job_No,T06Ref_No,T06Date,T06Time,T06Emp,T06Department,T06V_no,T06Cus_No,T06Status)" & _
                                                         " values('" & Trim(cboJob.Text) & "','" & Trim(txtEntry.Text) & "', '" & Trim(txtDate.Text) & "','" & Now & "','" & _Emp & "','" & Trim(cboDepartment.Text) & "','" & Trim(cboV_no.Text) & "','" & _CusNo & "','ISSUE')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "Insert Into tmpTransaction_Log(tmpStatus,tmpProcess,tmpTime,tmpUser,tmpCode)" & _
                                                    " values('ITEM ISSUE','SAVE', '" & Now & "','" & strDisname & "','" & Trim(txtEntry.Text) & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                'SAVE ISSUE FLUTER
                I = 0
                For Each uRow As UltraGridRow In UltraGrid1.Rows
                    nvcFieldList1 = "SELECT * FROM M05Item_Master WHERE M05Ref_No ='" & Trim(UltraGrid1.Rows(I).Cells(0).Text) & "'"
                    M02 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                    If isValidDataset(M02) Then
                        nvcFieldList1 = "Insert Into S01Stock_Balance(S01Item_Code,S01Ref_No,S01Date,S01Time,S01Tr_Type,S01Qty,S01Status)" & _
                                                " values('" & Trim(UltraGrid1.Rows(I).Cells(0).Text) & "','" & txtEntry.Text & "', '" & txtDate.Text & "','" & Now & "','ISSUE_ITEM','" & -(CDbl(UltraGrid1.Rows(I).Cells(5).Text) + CDbl(UltraGrid1.Rows(I).Cells(6).Text)) & "','A')"
                        ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                    End If
                    nvcFieldList1 = "Insert Into T07Item_Issue_Fluter(T07Ref_No,T07Item_Code,T07Part_No,T07Item_Name,T07Rate,T07Discount,T07Qty,T07Free,T07Status)" & _
                                          " values('" & txtEntry.Text & "','" & Trim(UltraGrid1.Rows(I).Cells(0).Text) & "','" & Trim(UltraGrid1.Rows(I).Cells(1).Text) & "','" & Trim(UltraGrid1.Rows(I).Cells(2).Text) & "', '" & Trim(UltraGrid1.Rows(I).Cells(3).Text) & "','" & Trim(UltraGrid1.Rows(I).Cells(4).Text) & "','" & Trim(UltraGrid1.Rows(I).Cells(5).Text) & "','" & Trim(UltraGrid1.Rows(I).Cells(6).Text) & "','A')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                    I = I + 1
                Next
                MsgBox("Items issue successfully", MsgBoxStyle.Information, "Information ........")
               

            End If
            transaction.Commit()
            connection.ClearAllPools()
            connection.Close()
            Call Clear_Text()
            cboJob.ToggleDropdown()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.ClearAllPools()
                connection.Close()
            End If
        End Try
    End Function
    Function SEARCH_RECORDS()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Dim Value As Double
        Dim I As Integer
        Dim _st As String

        Try
            Sql = "select * from T06Item_Issue_Header INNER JOIN M10Employee ON M10Code=T06Emp where  T06Ref_No='" & Trim(txtEntry.Text) & "' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then

                cboJob.Text = Trim(M01.Tables(0).Rows(0)("T06Job_No"))
                cboDepartment.Text = Trim(M01.Tables(0).Rows(0)("T06Department"))
                txtDate.Text = Trim(M01.Tables(0).Rows(0)("T06Date"))
                cboEmp.Text = Trim(M01.Tables(0).Rows(0)("M10Name"))
                cmdDelete.Enabled = True
                Call Search_Jobno()
               
            End If
            I = 0
            Call Load_Gride2()
            Sql = "SELECT * FROM T07Item_Issue_Fluter WHERE T07Ref_No='" & Trim(txtEntry.Text) & "' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            For Each DTRow2 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow
                newRow("#Ref.No") = Trim(M01.Tables(0).Rows(I)("T07Item_Code"))
                newRow("#Part No") = Trim(M01.Tables(0).Rows(I)("T07Part_No"))
                newRow("Item Name") = Trim(M01.Tables(0).Rows(I)("T07Item_Name"))
                Value = Trim(M01.Tables(0).Rows(I)("T07Rate"))
                _st = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _st = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Retail Price") = _st
                newRow("Qty") = Trim(M01.Tables(0).Rows(I)("T07Qty"))
                newRow("Free Issue") = Trim(M01.Tables(0).Rows(I)("T07Free"))
                newRow("Discount%") = Trim(M01.Tables(0).Rows(I)("T07Discount"))
                Value = CDbl(M01.Tables(0).Rows(I)("T07Rate")) * CDbl(M01.Tables(0).Rows(I)("T07Qty"))
                Value = Value - (CDbl(M01.Tables(0).Rows(I)("T07Rate")) * CDbl(M01.Tables(0).Rows(I)("T07Qty"))) * CDbl(M01.Tables(0).Rows(I)("T07Discount")) / 100
                _st = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _st = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Total") = _st
                c_dataCustomer1.Rows.Add(newRow)
                I = I + 1
            Next
            txtCount.Text = UltraGrid1.Rows.Count


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
    Function Clear_Text()
        Call Load_Item()
        Call Load_Gride2()
        Call Load_EntryNo()
        Call Load_JobNo()
        Call Load_VNO()
        Me.txtRate.Text = ""
        Me.txtQty.Text = ""
        Me.txtDiscount.Text = "0"
        Me.txtTotal.Text = ""
        Me.txtMtr.Text = ""
        Me.cboEmp.Text = ""
        Me.cboJob.Text = ""
        Me.cboDepartment.Text = ""
        Me.cboV_no.Text = ""
        Me.txtCount.Text = ""
        Me.txtCustomer.Text = ""
        Me.txtAddress.Text = ""
        Me.txtTp.Text = ""
        Me.cboPartno.Text = ""
        Me.txtItem.Text = ""
        Me.txtRate.Text = ""
        Me.txtQty.Text = ""
        Me.txtTotal.Text = ""
        Me.txtDiscount.Text = ""
        Me.txtFree.Text = ""
        strItem_Code = ""
        UltraGrid3.Visible = False
        Me.cmdDelete.Enabled = False
        UltraGrid2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
    End Function

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        Call Clear_Text()
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
            UltraGrid3.DataSource = M01
            UltraGrid3.Rows.Band.Columns(0).Width = 40

            UltraGrid3.Rows.Band.Columns(1).Width = 90
            UltraGrid3.Rows.Band.Columns(2).Width = 110
            UltraGrid3.Rows.Band.Columns(3).Width = 210
            UltraGrid3.Rows.Band.Columns(4).Width = 80
            UltraGrid3.Rows.Band.Columns(5).Width = 80
            UltraGrid3.Rows.Band.Columns(6).Width = 80
            UltraGrid3.Rows.Band.Columns(7).Width = 80
            '  UltraGrid1.Rows.Band.Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            UltraGrid3.Rows.Band.Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid3.Rows.Band.Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid3.Rows.Band.Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid3.Rows.Band.Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
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
    Private Sub txtItem_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtItem.KeyUp
        If e.KeyCode = 13 Then
            txtRate.Focus()
        End If
    End Sub

    Private Sub ItemLookupToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ItemLookupToolStripMenuItem.Click
        strWindowName = Me.Name
        frmView_Item_Uniq.Close()
        frmView_Item_Uniq.Show()
    End Sub

    Private Sub DeactivateToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeactivateToolStripMenuItem.Click
        frmIssue_Note.Close()
        frmIssue_Note.Show()
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
        Dim _GetDate As DateTime
        Dim _Get_Time As DateTime
        Dim A As String

        Try
            A = MsgBox("Are you sure you want to cancel this issue note", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Cancel Issue Note ..........")
            If A = vbYes Then
                nvcFieldList1 = "select * from T05Job_Card where T05Job_No='" & Trim(cboJob.Text) & "' and T05Status='CLOSE'"
                MB51 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                If isValidDataset(MB51) Then
                    MsgBox("This Job No alrady close", MsgBoxStyle.Information, "Information ......")
                    connection.Close()
                    Exit Sub
                Else
                    nvcFieldList1 = "UPDATE T06Item_Issue_Header SET T06Status='CANCEL' WHERE T06Ref_No='" & txtEntry.Text & "'"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                    nvcFieldList1 = "UPDATE T07Item_Issue_Fluter SET T07Status='CANCEL' WHERE T07Ref_No='" & txtEntry.Text & "'"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                    nvcFieldList1 = "UPDATE S01Stock_Balance SET S01Status='CANCEL' WHERE S01Ref_No='" & txtEntry.Text & "' AND S01Tr_Type='ISSUE_ITEM'"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                    nvcFieldList1 = "Insert Into tmpTransaction_Log(tmpStatus,tmpProcess,tmpTime,tmpUser,tmpCode)" & _
                                                  " values('ITEM ISSUE','CANCEL', '" & Now & "','" & strDisname & "','" & Trim(txtEntry.Text) & "')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                    MsgBox("Item Issue Note cancel successfully", MsgBoxStyle.Information, "Information .......")
                End If
            End If
            transaction.Commit()
            connection.ClearAllPools()
            connection.Close()
            Call Clear_Text()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.ClearAllPools()
                connection.Close()
            End If
        End Try
    End Sub

    Private Sub cboJob1_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboJob1.AfterCloseUp
        Call Search_Jobno_1()
    End Sub

    Private Sub cboJob1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboJob1.KeyUp
        If e.KeyCode = 13 Then
            cboDepartment1.ToggleDropdown()
        End If
    End Sub

    Private Sub cboVno1_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboVno1.AfterCloseUp
        Call Search_Vehicle_No_1()
    End Sub

    Private Sub cboVno1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboVno1.KeyUp
        If e.KeyCode = 13 Then
            cboDepartment1.ToggleDropdown()
        End If
    End Sub

    Private Sub cboDepartment1_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboDepartment1.AfterCloseUp
        Call Load_Searvice()
    End Sub

    Private Sub cboDepartment1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboDepartment1.KeyUp
        If e.KeyCode = 13 Then
            cboService1.ToggleDropdown()
        ElseIf e.KeyCode = Keys.F1 Then
            OPRUser.Visible = True
            txtUserName.Text = ""
            txtPassword.Text = ""
            txtUserName.Focus()
        End If
    End Sub

    Private Sub cboService1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboService1.KeyUp
        If e.KeyCode = 13 Then
            txtRate1.Focus()
        End If
    End Sub

    Private Sub txtRate1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtRate1.KeyUp
        Dim Value As Double
        Dim _st As String
        Dim A As String
        Try
            If e.KeyCode = 13 Then
                If Trim(cboDepartment1.Text) <> "" Then
                Else
                    MsgBox("Please select the department", MsgBoxStyle.Information, "Information .......")
                    Exit Sub
                End If

                If Trim(cboService1.Text) <> "" Then
                Else
                    MsgBox("Please enter the technical description", MsgBoxStyle.Information, "Information .......")
                    Exit Sub
                End If

                If txtRate1.Text <> "" Then
                Else
                    MsgBox("Please enter the Technical Rate", MsgBoxStyle.Information, "Information ........")
                    Exit Sub
                End If

                If IsNumeric(txtRate1.Text) Then
                Else
                    MsgBox("Please enter the correct Technical Rate", MsgBoxStyle.Information, "Information ........")
                    Exit Sub
                End If

                Dim newRow As DataRow = c_dataCustomer2.NewRow
                newRow("Department") = Trim(cboDepartment1.Text)
                newRow("#Part No") = "-"
                newRow("Item Name") = UCase(Trim(cboService1.Text))
                Value = Trim(txtRate1.Text)
                _st = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _st = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Retail Price") = _st
                newRow("Qty") = "1"
                newRow("Free Issue") = "0"
                newRow("Discount%") = "0"
                Value = txtRate1.Text
                '   Value = Value - (CDbl(M01.Tables(0).Rows(I)("T07Rate")) * CDbl(M01.Tables(0).Rows(I)("T07Qty"))) * CDbl(M01.Tables(0).Rows(I)("T07Discount")) / 100
                _st = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _st = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Total") = _st
                c_dataCustomer2.Rows.Add(newRow)
                txtRate1.Text = ""
                cboService1.Text = ""
                cboDepartment1.ToggleDropdown()
                Call Calculation_Net()
            End If
            txtCount1.Text = UltraGrid2.Rows.Count
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                ' connection.Close()
            End If
        End Try
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
    Function Clear_Text1()
        Call Load_EntryNo1()
        Call Load_Gride3()
        Me.cboJob1.Text = ""
        Me.cboVno1.Text = ""
        Me.txtMtr1.Text = ""
        Me.cboService1.Text = ""
        Me.txtNet1.Text = ""
        Me.cboDepartment1.Text = ""
        Me.cboService1.Text = ""
        Me.txtNext.Text = ""
        Me.cboService.Text = ""
        Me.txtRate1.Text = ""
        Me.txtCount1.Text = ""
        Me.txtCname1.Text = ""
        Me.txtAddress1.Text = ""
        Me.txtTP1.Text = ""
        Me.cboEmp01.Text = ""
        Me.cboEmp02.Text = ""
        Me.cboService.Text = ""
        Me.ToolStripMenuItem2.Enabled = False
        Me.cmdDelete1.Enabled = False
        OPRUser.Visible = False
        _LogStaus = False
        Call Load_JobNo()
        Call Load_VNO()
        UltraGrid2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
    End Function

    Private Sub UltraButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton3.Click
        Call Clear_Text1()
    End Sub

    Private Sub cboService_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboService.AfterCloseUp
        If IsNumeric(cboService.Text) Then
            If IsNumeric(txtMtr1.Text) Then
                txtNext.Text = CDbl(txtMtr1.Text) + CDbl(cboService.Text)

            End If
        Else
            txtNext.Text = "0"
        End If
        
    End Sub

    Private Sub cboService_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboService.KeyUp
        If e.KeyCode = 13 Then
            cboEmp01.ToggleDropdown()
        End If
    End Sub


    Private Sub cboService_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboService.TextChanged
        If IsNumeric(cboService.Text) Then
            If IsNumeric(txtMtr1.Text) Then
                txtNext.Text = CDbl(txtMtr1.Text) + CDbl(cboService.Text)

            End If
        Else
            txtNext.Text = "0"
        End If

       
    End Sub

    Private Sub cboEmp01_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboEmp01.KeyUp
        If e.KeyCode = 13 Then
            cboEmp02.ToggleDropdown()
        End If
    End Sub

    Private Sub cboEmp02_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboEmp02.KeyUp
        If e.KeyCode = 13 Then
            cboDepartment1.ToggleDropdown()
        End If
    End Sub

    Private Sub UltraButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton2.Click
        'If Search_Jobno_1() = True Then
        'Else
        '    MsgBox("Please select the correct Job No", MsgBoxStyle.Information, "Information ......")
        '    cboJob1.ToggleDropdown()
        '    Exit Sub
        'End If

        If Trim(cboEmp01.Text) <> "" Then
            If Search_Emp_1() = True Then
            Else
                MsgBox("Please select the Technicion", MsgBoxStyle.Information, "Information ........")
                cboEmp01.ToggleDropdown()
                Exit Sub
            End If
        End If

        If Trim(cboEmp02.Text) <> "" Then
            If Search_Emp_2() = True Then
            Else
                MsgBox("Please select the Technicion", MsgBoxStyle.Information, "Information ........")
                cboEmp02.ToggleDropdown()
                Exit Sub
            End If
        End If

        If UltraGrid2.Rows.Count > 0 Then
        Else
            MsgBox("Please enter the Labour chargers", MsgBoxStyle.Information, "Information ....")
            Exit Sub
        End If

        If Trim(cboEmp01.Text) <> "" And Trim(cboEmp02.Text) <> "" Then
            If Trim(cboEmp01.Text) = Trim(cboEmp02.Text) Then
                MsgBox("Technicion01 and Technocion02 same,Please change the technicion01 or technicion02", MsgBoxStyle.Information, "Information .....")
                Exit Sub
            End If
        End If

        frmPayMain_uniq.Close()
        frmPayMain_uniq.Show()
        frmPayMain_uniq.txtBill_Amount.Text = txtNet1.Text
        frmPayMain_uniq.txtBalance.Text = txtNet1.Text
        '  Call Save_Invoice()
    End Sub

    Private Sub UltraGrid2_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles UltraGrid2.AfterRowsDeleted
        Call Calculation_Net()
    End Sub

    Function Save_Invoice()
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
        Dim M02 As DataSet
        Dim M01 As DataSet
        Dim i As Integer
        Dim _Cost As Double
        Dim _SERVICE_AMOUNT As Double
        Dim _REMARK As String

        ' Dim _Gettime As Date

        '   Dim A As String
        Try
            _GetDate = Month(txtDate1.Text) & "/" & Microsoft.VisualBasic.Day(txtDate1.Text) & "/" & Year(txtDate1.Text)

            _Get_Time = Month(Now) & "/" & Microsoft.VisualBasic.Day(Now) & "/" & Year(Now) & " " & Hour(Now) & ":" & Minute(Now)
            nvcFieldList1 = "SELECT * FROM T08Sales_Header WHERE T08Invo_No='" & Trim(txtEnter1.Text) & "' "
            MB51 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(MB51) Then
                MsgBox("Can't change this invoice", MsgBoxStyle.Information, "Information ..........")
                connection.Close()
                Exit Function
            Else
                Call Load_EntryNo1()

                nvcFieldList1 = "update P01Parameter set P01No=P01No+ " & 1 & " where P01Code='INV' "
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                i = 0
                _SERVICE_AMOUNT = 0
                For Each uRow As UltraGridRow In UltraGrid2.Rows
                    If Trim(UltraGrid2.Rows(i).Cells(0).Text) = "-" Then
                    Else
                        'nvcFieldList1 = "Insert Into T09Sales_Flutter(T09Inv_No,T09Department,T09Item_Code,T09Item_Name,T09Cost,T09Retail,T09Qty,T09Discount,T09Free,T09Status)" & _
                        '                          " values('" & Trim(txtEnter1.Text) & "','" & Trim(UltraGrid2.Rows(i).Cells(0).Text) & "','-','" & Trim(UltraGrid2.Rows(i).Cells(1).Text) & "','0','" & Trim(UltraGrid2.Rows(i).Cells(2).Text) & "','" & Trim(UltraGrid2.Rows(i).Cells(4).Text) & "','0','0','A')"
                        'ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                        _SERVICE_AMOUNT = _SERVICE_AMOUNT + CDbl(Trim(UltraGrid2.Rows(i).Cells(7).Text))
                    End If

                    i = i + 1
                Next
                '===============================================================================

                'COMMISION PAY

                nvcFieldList1 = "SELECT * FROM T10Technicion_Comm WHERE T10INV_No='" & Trim(txtEnter1.Text) & "' AND T10Job_No='" & Trim(cboJob1.Text) & "'"
                M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                If isValidDataset(M01) Then
                    'nvcFieldList1 = "select * from M10Employee where M10Name='" & Trim(cboEmp01.Text) & "' and M10Status='A'"
                    'M02 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                    'If isValidDataset(M02) Then
                    '    nvcFieldList1 = "UPDATE T10Technicion_Comm SET T10Amount='" & _SERVICE_AMOUNT * (CDbl(M02.Tables(0).Rows(0)("M10Service_Cg")) / 100) & "' where T10INV_No='" & Trim(txtEnter1.Text) & "' and T10Emp='" & Trim(M02.Tables(0).Rows(0)("M10Code")) & "' and T10Status='A'"
                    '    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                    'End If

                    'nvcFieldList1 = "select * from M10Employee where M10Name='" & Trim(cboEmp02.Text) & "' and M10Status='A'"
                    'M02 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                    'If isValidDataset(M02) Then
                    '    nvcFieldList1 = "UPDATE T10Technicion_Comm SET T10Amount='" & _SERVICE_AMOUNT * (CDbl(M02.Tables(0).Rows(0)("M10Service_Cg")) / 100) & "' where T10INV_No='" & Trim(txtEnter1.Text) & "' and T10Emp='" & Trim(M02.Tables(0).Rows(0)("M10Code")) & "' and T10Status='A'"
                    '    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                    'End If
                Else
                    If Trim(cboEmp01.Text) <> "" Then
                        nvcFieldList1 = "select * from M10Employee where M10Name='" & Trim(cboEmp01.Text) & "' and M10Status='A'"
                        M02 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                        If isValidDataset(M02) Then
                            nvcFieldList1 = "Insert Into T10Technicion_Comm(T10INV_No,T10Job_No,T10Emp,T10Com_Rate,T10Amount,T10Pay_No,T10Status)" & _
                                                    " values('" & Trim(txtEnter1.Text) & "','" & Trim(cboJob1.Text) & "','" & Trim(M02.Tables(0).Rows(0)("M10Code")) & "','" & Trim(M02.Tables(0).Rows(0)("M10Service_Cg")) & "','" & _SERVICE_AMOUNT * (CDbl(M02.Tables(0).Rows(0)("M10Service_Cg")) / 100) & "','-','A')"
                            ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                        End If
                    End If

                    If Trim(cboEmp02.Text) <> "" Then
                        nvcFieldList1 = "select * from M10Employee where M10Name='" & Trim(cboEmp02.Text) & "' and M10Status='A'"
                        M02 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                        If isValidDataset(M02) Then
                            nvcFieldList1 = "Insert Into T10Technicion_Comm(T10INV_No,T10Job_No,T10Emp,T10Com_Rate,T10Amount,T10Pay_No,T10Status)" & _
                                                    " values('" & Trim(txtEnter1.Text) & "','" & Trim(cboJob1.Text) & "','" & Trim(M02.Tables(0).Rows(0)("M10Code")) & "','" & Trim(M02.Tables(0).Rows(0)("M10Service_Cg")) & "','" & _SERVICE_AMOUNT * (CDbl(M02.Tables(0).Rows(0)("M10Service_Cg")) / 100) & "','-','A')"
                            ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                        End If
                    End If
                End If
                '=============================================================================================================================

                'SAVE INVOICE 
                nvcFieldList1 = "Insert Into T08Sales_Header(T08Invo_No,T08Job_No,T08Tr_Type,T08Service_on,T08St_Mtr,T08End_mtr,T08Date,T08V_No,T08Cus_NO,T08Status,T08Net_Amount,T08Time)" & _
                                                         " values('" & Trim(txtEnter1.Text) & "','" & Trim(cboJob1.Text) & "', 'JOB_INVOICE','" & Trim(cboService.Text) & "','" & Trim(txtMtr1.Text) & "','" & Trim(txtNext.Text) & "','" & _GetDate & "','" & Trim(cboVno1.Text) & "','" & _CusNo & "','A','" & CDbl(txtNet1.Text) & "','" & _Get_Time & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "Insert Into tmpTransaction_Log(tmpStatus,tmpProcess,tmpTime,tmpUser,tmpCode)" & _
                                                    " values('JOB_INVOICE','SAVE', '" & _Get_Time & "','" & strDisname & "','" & Trim(txtEnter1.Text) & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                '=======================================================================
                'UPDATE JOB CARD HEADER
                nvcFieldList1 = "update T05Job_Card set T05Status='CLOSE' where T05Job_No='" & Trim(cboJob1.Text) & "' "
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                '===============================================================================================
                nvcFieldList1 = "select * from T07Item_Issue_Fluter inner join T06Item_Issue_Header on T06Ref_No=T07Ref_No where T06Job_No='" & Trim(cboJob1.Text) & "' and T07Status='A'"
                M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                i = 0
                For Each DTRow2 As DataRow In M01.Tables(0).Rows
                    _Cost = 0
                    If Trim(UltraGrid2.Rows(i).Cells(0).Text) = "-" Then
                        nvcFieldList1 = "select * from M05Item_Master where M05Item_Code='" & Trim(M01.Tables(0).Rows(i)("T07Part_No")) & "'"
                        M02 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                        If isValidDataset(M02) Then
                            _Cost = Trim(M02.Tables(0).Rows(0)("M05Cost"))
                        End If

                        nvcFieldList1 = "Insert Into T09Sales_Flutter(T09Inv_No,T09Department,T09Item_Code,T09Item_Name,T09Cost,T09Retail,T09Qty,T09Discount,T09Free,T09Status)" & _
                                                          " values('" & Trim(txtEnter1.Text) & "','-','" & Trim(M01.Tables(0).Rows(i)("T07Part_No")) & "','" & Trim(M01.Tables(0).Rows(i)("T07Item_Name")) & "','" & _Cost & "','" & Trim(M01.Tables(0).Rows(i)("T07Rate")) & "','" & Trim(M01.Tables(0).Rows(i)("T07Qty")) & "','" & Trim(M01.Tables(0).Rows(i)("T07Discount")) & "','" & Trim(M01.Tables(0).Rows(i)("T07Free")) & "','A')"
                        ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                    End If
                    i = i + 1
                Next

                i = 0
                '  _SERVICE_AMOUNT = 0
                For Each uRow As UltraGridRow In UltraGrid2.Rows
                    If Trim(UltraGrid2.Rows(i).Cells(0).Text) = "-" Then
                    Else
                        nvcFieldList1 = "Insert Into T09Sales_Flutter(T09Inv_No,T09Department,T09Item_Code,T09Item_Name,T09Cost,T09Retail,T09Qty,T09Discount,T09Free,T09Status)" & _
                                                  " values('" & Trim(txtEnter1.Text) & "','" & Trim(UltraGrid2.Rows(i).Cells(0).Text) & "','" & Trim(UltraGrid2.Rows(i).Cells(1).Text) & "','" & Trim(UltraGrid2.Rows(i).Cells(2).Text) & "','0','" & Trim(UltraGrid2.Rows(i).Cells(3).Text) & "','" & Trim(UltraGrid2.Rows(i).Cells(5).Text) & "','0','0','A')"
                        ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                        '  _SERVICE_AMOUNT = _SERVICE_AMOUNT + CDbl(Trim(UltraGrid2.Rows(i).Cells(7).Text))
                    End If

                    i = i + 1
                Next



                '=================================================================================================
                If CDbl(frmPayMain_uniq.txtBalance.Text) < 0 Then
                    frmPayMain_uniq.txtBalance.Text = "0"
                End If
                'PAY HEADER
                nvcFieldList1 = "select * from T11Income_Summery where T11Invo_No='" & Trim(txtEnter1.Text) & "'"
                M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                If isValidDataset(M01) Then
                Else
                    nvcFieldList1 = "Insert Into T11Income_Summery(T11Invo_No,T11Job_No,T11Tr_Type,T11Date,T11Net_Amount,T11Cash,T11Chq,T11Card,T11Credit,T11Status)" & _
                                                   " values('" & Trim(txtEnter1.Text) & "','" & Trim(cboJob1.Text) & "','JOB_INVOICE','" & _GetDate & "','" & CDbl(txtNet1.Text) & "','" & CDbl(frmPayMain_uniq.txtCash.Text) & "','" & CDbl(frmPayMain_uniq.txtChq_Total.Text) & "','" & CDbl(frmPayMain_uniq.txtTotal.Text) & "','" & CDbl(frmPayMain_uniq.txtBalance.Text) & "','A')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                End If
                '===================================================================================================
                'OUTSTANDING 
                If CDbl(frmPayMain_uniq.txtBalance.Text) > 0 Then
                    nvcFieldList1 = "select * from T12OutStanding where T12Inv_No='" & Trim(txtEnter1.Text) & "'"
                    M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                    If isValidDataset(M01) Then
                    Else
                        nvcFieldList1 = "Insert Into T12OutStanding(T12Inv_No,T12Cus_No,T12Date,T12Pay_No,T12Chq_No,T12Cr,T12Dr,T12Status)" & _
                                                       " values('" & Trim(txtEnter1.Text) & "','" & _CusNo & "','" & _GetDate & "','-','-','" & CDbl(frmPayMain_uniq.txtBalance.Text) & "','0','A')"
                        ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                    End If
                End If
                '===================================================================================================
                'CHQ TRANSACTION
                i = 0
                With frmPayMain_uniq
                    For Each uRow As UltraGridRow In .UltraGrid2.Rows
                        nvcFieldList1 = "Insert Into T13Chq_Transaction(T13Ref_No,T13Cus_Code,T13Date,T13Bank,T13Chq_No,T13DOR,T13Amount,T13Return_Status,T13Status,T13Tr_Type)" & _
                                                      " values('" & Trim(txtEnter1.Text) & "','" & _CusNo & "','" & _GetDate & "','" & .UltraGrid2.Rows(i).Cells(0).Text & "','" & .UltraGrid2.Rows(i).Cells(1).Text & "','" & .UltraGrid2.Rows(i).Cells(2).Text & "','" & .UltraGrid2.Rows(i).Cells(3).Text & "','-','A','JOB_INVOICE')"
                        ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                        i = i + 1
                    Next
                End With
                '===================================================================================================
                'CREDIT CARD TRANSACTION
                i = 0
                With frmPayMain_uniq
                    For Each uRow As UltraGridRow In .UltraGrid1.Rows
                        nvcFieldList1 = "Insert Into T14Credit_Card_TR(T14Inv_No,T14Date,T14Type,T14Card_No,T14Amount,T14Status)" & _
                                                      " values('" & Trim(txtEnter1.Text) & "','" & _GetDate & "','" & .UltraGrid1.Rows(i).Cells(0).Text & "','" & .UltraGrid1.Rows(i).Cells(1).Text & "','" & .UltraGrid1.Rows(i).Cells(2).Text & "','A')"
                        ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                        i = i + 1
                    Next
                End With
                '==================================================================================================
                'DAILY INCOME
                _SERVICE_AMOUNT = CDbl(frmPayMain_uniq.txtCash.Text) + CDbl(frmPayMain_uniq.txtTotal.Text)

                If _SERVICE_AMOUNT > 0 Then
                    _REMARK = "Workshop Income -" & Trim(txtEnter1.Text)

                    nvcFieldList1 = "Insert Into T04Profit_Loss(T04Date,T04Tr_Type,T04Ref_No,T04Description,T04Cr,T04Dr,T04Status)" & _
                                                     " values('" & _GetDate & "','JOB_INCOME','" & txtEnter1.Text & "','" & _REMARK & "','" & _SERVICE_AMOUNT & "','0','A')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                End If
            End If
            transaction.Commit()
            connection.ClearAllPools()
            connection.Close()
            A = MsgBox("Are you sure you want to print this Invoice", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Information .........")
            If A = vbYes Then

            End If
            Call Clear_Text1()
            cboJob1.ToggleDropdown()

        Catch returnMessage As Exception
            transaction.Rollback()
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.ClearAllPools()
                connection.Close()
            End If
        End Try
    End Function

    Private Sub cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        OPRUser.Visible = False
        _permissionLevel = ""
    End Sub

    Private Sub txtUserName_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = 13 Then
            If Trim(txtUserName.Text) <> "" Then
                txtPassword.Focus()
            End If
        End If
    End Sub

   

    Private Sub txtPassword_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
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
                _permissionLevel = Trim(M01.Tables(0).Rows(0)("UType"))
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

    Private Sub UltraTabControl1_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles UltraTabControl1.SelectedTabChanged
        Call Clear_Text1()
    End Sub

    Private Sub ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem1.Click
        frmView_Invoice_Job.Close()
        frmView_Invoice_Job.Show()
    End Sub

    Private Sub cmdDelete1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelete1.Click
        Dim A As String

        A = MsgBox("Are you sure you want to delete this invoice", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Information ..........")
        If A = vbYes Then
            If _LogStaus = True Then
                Call Delete_Invoice()
            Else
                OPRUser.Visible = True
                txtUserName.Focus()
            End If
        End If
    End Sub

    Function Delete_Invoice()
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
        Dim M01 As DataSet
        Try
            nvcFieldList1 = "select * from T05Job_Card  where T05Job_No='" & Trim(cboJob1.Text) & "' and T05Status='CLOSE' "
            M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(M01) Then
                nvcFieldList1 = "update T05Job_Card set T05Status='A' where T05Job_No='" & Trim(cboJob1.Text) & "' "
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "update T08Sales_Header set T08Status='CANCEL' where T08Invo_No='" & Trim(txtEnter1.Text) & "' "
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "update T09Sales_Flutter set T09Status='CANCEL' where T09Inv_No='" & Trim(txtEnter1.Text) & "' "
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                'TECHNICIAN COMMISSION 
                nvcFieldList1 = "SELECT * FROM T10Technicion_Comm WHERE T10INV_No='" & txtEnter1.Text & "' AND T10Status='PAID'"
                MB51 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                If isValidDataset(MB51) Then
                    MsgBox("Can't delete this invoice technician commission alrady paid", MsgBoxStyle.Information, "Information .......")
                    connection.Close()
                    Exit Function
                
                Else
                    nvcFieldList1 = "update T10Technicion_Comm set T10Status='CANCEL' where T10INV_No='" & Trim(txtEnter1.Text) & "' "
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                End If
                '==================================================================================================
                'OUTSTANDING PAY
                nvcFieldList1 = "SELECT * FROM T15Outstanding_Collection WHERE T15Inv_No='" & txtEnter1.Text & "' AND T15Status='PAID'"
                MB51 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                If isValidDataset(MB51) Then
                    MsgBox("Can't delete this invoice Customer alrady paid", MsgBoxStyle.Information, "Information .......")
                    connection.Close()
                    Exit Function
                    
                Else
                    nvcFieldList1 = "update T12OutStanding set T12Status='CANCEL' where T12Inv_No='" & Trim(txtEnter1.Text) & "' "
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                End If
                '====================================================================================================
                'CHQ TRANSACTION
                nvcFieldList1 = "update T13Chq_Transaction set T13Status='CANCEL' where T13Ref_No='" & Trim(txtEnter1.Text) & "' "
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                '====================================================================================================
                'CREDIT CARD
                nvcFieldList1 = "update T14Credit_Card_TR set T14Status='CANCEL' where T14Inv_No='" & Trim(txtEnter1.Text) & "' "
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                '====================================================================================================
                'INCOME STATMENT
                nvcFieldList1 = "update T11Income_Summery set T11Status='CANCEL' where T11Invo_No='" & Trim(txtEnter1.Text) & "' AND T11Tr_Type='JOB_INVOICE' "
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                '=====================================================================================================
                'P&L REPORT
                nvcFieldList1 = "update T04Profit_Loss set T04Status='CANCEL' where T04Ref_No='" & Trim(txtEnter1.Text) & "' AND T04Tr_Type='JOB_INCOME' "
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "Insert Into tmpTransaction_Log(tmpStatus,tmpProcess,tmpTime,tmpUser,tmpCode)" & _
                                             " values('JOB_INVOICE','DELETE', '" & _Get_Time & "','" & strDisname & "','" & Trim(txtEnter1.Text) & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)


                MsgBox("Invoice deleted successfully", MsgBoxStyle.Information, "Information ........")
                transaction.Commit()
                ' connection.Close()
                Call Clear_Text1()
            End If
            connection.ClearAllPools()
            connection.Close()
        Catch returnMessage As Exception
            transaction.Rollback()
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.ClearAllPools()
                connection.Close()
            End If

        End Try
    End Function

    Private Sub txtUserName_KeyUp1(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtUserName.KeyUp
        If e.KeyCode = 13 Then
            txtPassword.Focus()
        End If
    End Sub


    Private Sub RefreshToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RefreshToolStripMenuItem.Click
        Call Clear_Text()
        Call Clear_Text1()

    End Sub


    Private Sub UltraGrid2_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles UltraGrid2.DoubleClickRow
        _lastRow = UltraGrid2.ActiveRow.Index
    End Sub

   
    Private Sub UltraGrid2_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles UltraGrid2.KeyUp
        On Error Resume Next
        If e.KeyCode = Keys.Delete Then
            _lastRow = UltraGrid2.ActiveRow.Index
            If Trim(UltraGrid2.Rows(_lastRow).Cells(0).Text) = "-" Then
            Else
                UltraGrid2.Rows(_lastRow).Delete()
            End If
        End If

        txtCount1.Text = UltraGrid2.Rows.Count
    End Sub

   

    Private Sub cboDepartment1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboDepartment1.TextChanged
        Call Load_Searvice()
    End Sub

    Private Sub cboPartno_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboPartno.TextChanged
        If UltraGrid3.Visible = True Then
            Call Load_Grid_ROW()
        Else
            UltraGrid3.Visible = True
            Call Load_Grid_ROW()
        End If
    End Sub

    Private Sub UltraGrid3_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles UltraGrid3.DoubleClickRow
        Dim _Row As Integer

        _Row = UltraGrid3.ActiveRow.Index
        strItem_Code = Trim(UltraGrid3.Rows(_Row).Cells(0).Text)
        Call Search_Item()
        UltraGrid3.Visible = False
        txtQty.Focus()
    End Sub

    Private Sub UltraGrid3_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles UltraGrid3.KeyUp
        Dim _Row As Integer
        If e.KeyCode = 13 Then
            _Row = UltraGrid3.ActiveRow.Index
            strItem_Code = Trim(UltraGrid3.Rows(_Row).Cells(0).Text)
            Call Search_Item()
            UltraGrid3.Visible = False
            txtQty.Focus()
        End If
    End Sub
End Class