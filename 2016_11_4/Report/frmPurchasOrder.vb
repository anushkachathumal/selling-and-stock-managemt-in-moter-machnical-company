Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports CrystalDecisions.CrystalReports.Engine
Public Class frmPurchasOrder
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim dblInsuaranceCommision As Double
    Dim c_dataCustomer As DataTable
    Dim strPrice As Double
    Dim strTicket_price As Double
    Dim strSupplierscode As String
    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Private Sub frmPurchasOrder_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()

        txtAddress.ReadOnly = True
        txtCompanyname.ReadOnly = True
        txtItemName.ReadOnly = True
        txtPo.ReadOnly = True
        Loadgride()
        Try
            Sql = "select M03SuppCode as [Company Code],M03Name as [Description] from M03SuppliersDetailes"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            cboCompany.DataSource = dsUser
            cboCompany.Rows.Band.Columns(0).Width = 125
            cboCompany.Rows.Band.Columns(1).Width = 470
            '------------------------------------------------------------------------------

            Sql = "select M04ItemCode as [Item Code],M04ItemName as [Description] from M04Item where M04Status='A'"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            cboItem.DataSource = dsUser
            cboItem.Rows.Band.Columns(0).Width = 125
            cboItem.Rows.Band.Columns(1).Width = 470
            '------------------------------------------------------------------------------
            

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Sub

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Try
            Clicked = "ADD"
            OPR0.Enabled = True
            OPR1.Enabled = True
            OPR2.Enabled = True
            ' Call Clear_Text()
            cmdAdd.Enabled = False
            txtDate.Text = Today
            txtRecevingDate.Text = Today
            cboCompany.ToggleDropdown()

            Sql = "select * from P01parameter where P01code='PO'"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(dsUser) Then
                txtPo.Text = dsUser.Tables(0).Rows(0)("P01LastNo")
            End If

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Sub

    Function Loadgride()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer = CustomerDataClass.MakeDataTablePurchasorder
        UltraGrid1.DataSource = c_dataCustomer
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 120
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 260
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 110
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(3).Width = 110
            .DisplayLayout.Bands(0).Columns(3).AutoEdit = False
            '.DisplayLayout.Bands(0).Columns(4).Width = 90
            '.DisplayLayout.Bands(0).Columns(4).AutoEdit = False
            '.DisplayLayout.Bands(0).Columns(5).Width = 110
            '.DisplayLayout.Bands(0).Columns(5).AutoEdit = False
        End With
    End Function

    Function Searchcompany() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        Dim recCompany As DataSet
        con = DBEngin.GetConnection()

        Searchcompany = False
        Try
            Sql = "select * from M03SuppliersDetailes where M03SuppCode='" & Trim(cboCompany.Text) & "'"
            recCompany = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(recCompany) Then
                txtCompanyname.Text = recCompany.Tables(0).Rows(0)("M03Name")
                txtAddress.Text = recCompany.Tables(0).Rows(0)("M03Address")
                Searchcompany = True
            Else
                txtCompanyname.Text = ""
                Searchcompany = False
            End If
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub cboCompany_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboCompany.AfterCloseUp
        Searchcompany()
    End Sub

    Private Sub cboCompany_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles cboCompany.InitializeLayout

    End Sub

    Private Sub cboCompany_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboCompany.KeyUp
        If e.KeyCode = 13 Then
            Searchcompany()
            cboItem.ToggleDropdown()
        End If
    End Sub

    Function SearchItem() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        Dim recItemDetail As DataSet
        con = DBEngin.GetConnection()

        SearchItem = False
        Try
            Sql = "select * from M04Item where M04ItemCode='" & Trim(cboItem.Text) & "'"
            recItemDetail = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(recItemDetail) Then
                txtItemName.Text = recItemDetail.Tables(0).Rows(0)("M04ItemName")
                SearchItem = True
            Else
                SearchItem = False
                txtItemName.Text = ""
            End If

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub cboItem_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboItem.AfterCloseUp
        SearchItem()
    End Sub

    Private Sub cboItem_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles cboItem.InitializeLayout

    End Sub

    Private Sub cboItem_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboItem.KeyUp
        If e.KeyCode = 13 Then
            If cboItem.Text <> "" Then
                SearchItem()
                txtRecevingDate.Focus()
            Else
                cmdSave.Focus()
            End If
        End If
    End Sub

    Private Sub txtRecevingDate_BeforeDropDown(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtRecevingDate.BeforeDropDown

    End Sub

    Private Sub txtRecevingDate_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtRecevingDate.KeyUp
        If e.KeyCode = 13 Then
            txtQty.Focus()
        End If
    End Sub

    Private Sub txtQty_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtQty.KeyUp
        If e.KeyCode = 13 Then
            If Searchcompany() = True Then
                If SearchItem() = True Then
                    Dim newRow As DataRow = c_dataCustomer.NewRow
                    Try
                        newRow("Item Code") = cboItem.Text
                        newRow("Description") = txtItemName.Text
                        newRow("Date of Receving") = txtRecevingDate.Text
                        newRow("Quntity") = txtQty.Text
                        'newRow("Discount") = txtQty.Text
                        ' newRow("Total") = txtTotal.Text
                        c_dataCustomer.Rows.Add(newRow)

                        txtQty.Text = ""
                        txtItemName.Text = ""
                        cboItem.Text = ""
                        cboItem.ToggleDropdown()
                    Catch returnMessage As Exception
                        If returnMessage.Message <> Nothing Then
                            MessageBox.Show(returnMessage.Message)
                        End If
                    End Try
                Else
                    MsgBox("Please enter the correct Item code", MsgBoxStyle.Critical, "Sign Info ........")
                End If
            Else
                MsgBox("Please enter the correct Company", MsgBoxStyle.Information, "Sign Info .........")
            End If
        End If
    End Sub

    Private Sub txtQty_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtQty.TextChanged
        If IsNumeric(txtQty.Text) Then
            cmdSave.Enabled = True
        Else
            txtQty.Text = ""
        End If
    End Sub

    Private Sub txtQty_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtQty.ValueChanged

    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        Dim ncQryType As String
        Dim nvcFieldList As String
        Dim nvcWhereClause As String
        Dim nvcVccode As String
        Dim i As Integer

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True

        Dim B As New ReportDocument
        Dim A As String

        If Searchcompany() = True Then
            Try
                If UltraGrid1.Rows.Count > 0 Then
                    nvcFieldList = "Update P01parameter set P01LastNo = " & Val(txtPo.Text) + 1 & " where P01code = 'PO'"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList)

                    nvcFieldList = "Insert Into T10Purchaseorderheader(T10PONo,T10Date,T10companyCode)" & _
                                                                                         " values(" & Trim(txtPo.Text) & ",'" & txtDate.Text & "','" & cboCompany.Text & "')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList)

                    i = 0
                    For Each uRow As UltraGridRow In UltraGrid1.Rows
                        nvcFieldList = "Insert Into T11PurcheseorderFluter(T11PONo,T11ItemCode,T11Recevingdate,T11Qty)" & _
                                                                                                    " values(" & Trim(txtPo.Text) & ",'" & UltraGrid1.Rows(i).Cells(0).Value & "','" & UltraGrid1.Rows(i).Cells(2).Value & "','" & UltraGrid1.Rows(i).Cells(3).Value & "')"
                        ExecuteNonQueryText(connection, transaction, nvcFieldList)


                        i = i + 1
                    Next

                    transaction.Commit()

                    A = ConfigurationManager.AppSettings("ReportPath") + "\PO.rpt"

                    B.Load(A.ToString)
                    ' B.SetParameterValue("Todate", txtTo.Value)
                    'B.SetParameterValue("Fromdate", txtDate.Value)
                    frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                    frmReport.CrystalReportViewer1.DisplayToolbar = True
                    frmReport.CrystalReportViewer1.SelectionFormula = "{T10Purchaseorderheader.T10PONo}  =" & Trim(txtPo.Text) & ""
                    frmReport.MdiParent = MDIMain
                    'frmReport.Show()
                    frmReport.Show()
                    common.ClearAll(OPR0, OPR1, OPR2, OPR4)
                    Clicked = ""
                    cmdAdd.Enabled = True
                    cmdSave.Enabled = False
                    cmdEdit.Enabled = False
                    cmdDelete.Enabled = False
                    OPR4.Enabled = True
                    cmdAdd.Focus()
                    Call Loadgride()
                End If
                    Catch returnMessage As Exception
                If returnMessage.Message <> Nothing Then
                    MessageBox.Show(returnMessage.Message)
                End If
            End Try

        Else
                MsgBox("Please enter the company Code", MsgBoxStyle.Information, "Sign Info .......")
        End If
    End Sub
End Class