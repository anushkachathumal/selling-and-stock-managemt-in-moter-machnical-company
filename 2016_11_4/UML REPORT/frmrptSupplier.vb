Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports CrystalDecisions.CrystalReports.Engine
Public Class frmrptSupplier
    Dim c_dataCustomer1 As DataTable
    Dim _PrintStatus As String
    Dim _Comcode As String
    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Function Load_Grid()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M09Type as [##],M09Code as [Supplier Code],M09Name as [supplier Name],M09Address as [Address],M09TP as [Contact No] from M09Supplier where M09Active='A' and M09Loc_Code='" & _Comcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = M01
            With UltraGrid1
                .DisplayLayout.Bands(0).Columns(0).Width = 90
                .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
                .DisplayLayout.Bands(0).Columns(1).Width = 90
                .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
                .DisplayLayout.Bands(0).Columns(2).Width = 270
                .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
                .DisplayLayout.Bands(0).Columns(3).Width = 170
                .DisplayLayout.Bands(0).Columns(3).AutoEdit = False
                .DisplayLayout.Bands(0).Columns(4).Width = 90
                .DisplayLayout.Bands(0).Columns(4).AutoEdit = False
            End With
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Private Sub frmrptSupplier_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        _Comcode = ConfigurationManager.AppSettings("LOCCode")
        Call Load_Grid()
        txtAdd1.ReadOnly = True
        txtAddress.ReadOnly = True
        txtCode.ReadOnly = True
        txtContact.ReadOnly = True
        txtVAT.ReadOnly = True
        txtTp.ReadOnly = True
        txtFax.ReadOnly = True
        txtStatus.ReadOnly = True
        txtName.ReadOnly = True
        txtType.ReadOnly = True
        _PrintStatus = "A"
    End Sub

    Private Sub ActiveSupplierToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ActiveSupplierToolStripMenuItem.Click
        Call Load_Grid()
        _PrintStatus = "A"
        Call Load_Grid1()
    End Sub

    Private Sub GASToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ' Load_Grid1("VEGETABLE")
        _PrintStatus = "A1"
    End Sub

    Function Load_Grid1()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M09Type as [##],M09Code as [Supplier Code],M09Name as [supplier Name],M09Address as [Address],M09TP as [Contact No] from M09Supplier where M09Active='A'  and M09Loc_Code='" & _Comcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = M01
            With UltraGrid1
                .DisplayLayout.Bands(0).Columns(0).Width = 90
                .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
                .DisplayLayout.Bands(0).Columns(1).Width = 90
                .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
                .DisplayLayout.Bands(0).Columns(2).Width = 270
                .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
                .DisplayLayout.Bands(0).Columns(3).Width = 170
                .DisplayLayout.Bands(0).Columns(3).AutoEdit = False
                .DisplayLayout.Bands(0).Columns(4).Width = 90
                .DisplayLayout.Bands(0).Columns(4).AutoEdit = False
            End With
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Function Load_GridInactive()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M09Type as [##],M09Code as [Supplier Code],M09Name as [supplier Name],M09Address as [Address],M09TP as [Contact No] from M09Supplier where M09Active='I' and M09Loc_Code='" & _Comcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = M01
            With UltraGrid1
                .DisplayLayout.Bands(0).Columns(0).Width = 90
                .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
                .DisplayLayout.Bands(0).Columns(1).Width = 90
                .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
                .DisplayLayout.Bands(0).Columns(2).Width = 270
                .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
                .DisplayLayout.Bands(0).Columns(3).Width = 170
                .DisplayLayout.Bands(0).Columns(3).AutoEdit = False
                .DisplayLayout.Bands(0).Columns(4).Width = 90
                .DisplayLayout.Bands(0).Columns(4).AutoEdit = False
            End With
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Private Sub N2ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ' Load_Grid1("Leather")
        _PrintStatus = "A2"
    End Sub

    Private Sub SaftyItemsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ' Load_Grid1("Sole")
        _PrintStatus = "A3"
    End Sub

    Private Sub ChemicleToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '  Load_Grid1("Other")
        _PrintStatus = "A4"
    End Sub

    Private Sub OthersToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ' Load_Grid1("OTHER")
        _PrintStatus = "A8"
    End Sub

    Private Sub UltraGrid1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles UltraGrid1.DoubleClick
        On Error Resume Next
        Dim _RowIndex As Integer
        Dim _SupCode As String
        _RowIndex = UltraGrid1.ActiveRow.Index

        _SupCode = UltraGrid1.Rows(_RowIndex).Cells(1).Text
        Search_Record(_SupCode)
        OPR0.Visible = True
        Panel1.Visible = True
    End Sub


    Function Search_Record(ByVal strCode As String)
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select * from M09Supplier where M09Active='A' and M09Code='" & strCode & "' and M09Loc_Code='" & _Comcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                With M01
                    txtCode.Text = strCode
                    txtStatus.Text = .Tables(0).Rows(0)("M09Status")
                    txtName.Text = .Tables(0).Rows(0)("M09Name")
                    txtType.Text = .Tables(0).Rows(0)("M09Type")
                    txtAddress.Text = .Tables(0).Rows(0)("M09Address")
                    txtAdd1.Text = .Tables(0).Rows(0)("M09Address1")
                    txtContact.Text = .Tables(0).Rows(0)("M09Contact_On")
                    txtFax.Text = .Tables(0).Rows(0)("M09Fax")
                    txtTp.Text = .Tables(0).Rows(0)("M09TP")
                    txtVAT.Text = .Tables(0).Rows(0)("M09VAT")
                End With
            End If
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Private Sub ResetToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ResetToolStripMenuItem.Click
        OPR0.Visible = False
        Panel1.Visible = False
        Call Load_Grid()
    End Sub

    Private Sub PrintToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrintToolStripMenuItem.Click
        Dim B As New ReportDocument
        Dim A As String

        Try
            A = ConfigurationManager.AppSettings("ReportPath") + "\rptSupplire.rpt"
            B.Load(A.ToString)
            B.SetDatabaseLogon("sa", "tommya")

            '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
            frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
            frmReport.CrystalReportViewer1.DisplayToolbar = True
            If _PrintStatus = "A" Then
                frmReport.CrystalReportViewer1.SelectionFormula = "{M09Supplier.M09Active}='A' and {M09Supplier.M09Loc_Code}='" & _Comcode & "'"
            ElseIf _PrintStatus = "A1" Then
                frmReport.CrystalReportViewer1.SelectionFormula = "{M09Supplier.M09Active}='A' and {M09Supplier.M09Type}='VEGETABLE'"
            ElseIf _PrintStatus = "A2" Then
                frmReport.CrystalReportViewer1.SelectionFormula = "{M09Supplier.M09Active}='A' and {M09Supplier.M09Type}='Leather'"
            ElseIf _PrintStatus = "A3" Then
                frmReport.CrystalReportViewer1.SelectionFormula = "{M09Supplier.M09Active}='A' and {M09Supplier.M09Type}='Sole'"
            ElseIf _PrintStatus = "A4" Then
                frmReport.CrystalReportViewer1.SelectionFormula = "{M09Supplier.M09Active}='A' and {M09Supplier.M09Type}='OTHER'"
            
            ElseIf _PrintStatus = "A6" Then
                frmReport.CrystalReportViewer1.SelectionFormula = "{M09Supplier.M09Active}='I' and {M09Supplier.M09Loc_Code}='" & _Comcode & "' "
            End If
            frmReport.Refresh()
            ' frmReport.CrystalReportViewer1.PrintReport()
            ' B.PrintToPrinter(1, True, 0, 0)
            frmReport.MdiParent = MDIMain
            frmReport.Show()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                '   MsgBox(i)
            End If
        End Try
    End Sub

    

    Private Sub InactiveSupplierToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles InactiveSupplierToolStripMenuItem.Click
        _PrintStatus = "A6"
        Call Load_GridInactive()
    End Sub
End Class