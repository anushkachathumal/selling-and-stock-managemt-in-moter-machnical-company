Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports CrystalDecisions.CrystalReports.Engine
Public Class frmrptRowMaterial
    Dim c_dataCustomer1 As DataTable
    Dim _PrintStatus As String
    Dim _Category As String

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Function Load_Grid()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M12Category as [##],M12Item_Code as [Item Code],M12Name as [Item Name] from M12Row_Material   where M12Status='A' order by M12Category"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = M01
            With UltraGrid1
                .DisplayLayout.Bands(0).Columns(0).Width = 90
                .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
                .DisplayLayout.Bands(0).Columns(1).Width = 90
                .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
                .DisplayLayout.Bands(0).Columns(2).Width = 270
                .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
                
            End With
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Function Load_Grid1()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M12Category as [##],M12Item_Code as [Item Code],M12Name as [Item Name] from M12Row_Material   where M12Status='A' and M12Category='" & cboCategory.Text & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = M01
            With UltraGrid1
                .DisplayLayout.Bands(0).Columns(0).Width = 90
                .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
                .DisplayLayout.Bands(0).Columns(1).Width = 90
                .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
                .DisplayLayout.Bands(0).Columns(2).Width = 270
                .DisplayLayout.Bands(0).Columns(2).AutoEdit = False

            End With
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Private Sub frmrptRowMaterial_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_Grid()
        _PrintStatus = "All"

        Call Load_Category()
    End Sub

    Function Load_Category()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M08Name as [##] from M08Main_Category order by M08Code "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboCategory
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 140
                ' .Rows.Band.Columns(1).Width = 180


            End With

            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try

    End Function

    Private Sub AllItemsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AllItemsToolStripMenuItem.Click
        Call Load_Grid()
        _PrintStatus = "All"
        Panel1.Visible = False
        cboCategory.Text = ""
        _Category = ""
    End Sub

    Private Sub PrintToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrintToolStripMenuItem.Click
        Dim B As New ReportDocument
        Dim A As String

        Try
            A = ConfigurationManager.AppSettings("ReportPath") + "\rptRowmaterial.rpt"
            B.Load(A.ToString)
            B.SetDatabaseLogon("sa", "tommya")

            '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
            frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
            frmReport.CrystalReportViewer1.DisplayToolbar = True
            If _PrintStatus = "All" Then
                frmReport.CrystalReportViewer1.SelectionFormula = "{M12Row_Material.M12Status}='A'"
            ElseIf _PrintStatus = "Category" Then
                frmReport.CrystalReportViewer1.SelectionFormula = "{M12Row_Material.M12Status}='A' and {M12Row_Material.M12Category}='" & _Category & "'"
          
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

    Private Sub UltraButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton3.Click
        If cboCategory.Text <> "" Then
        Else
            MsgBox("Please select the category", MsgBoxStyle.Information, "Information .......")
            Exit Sub
        End If

        Call Load_Grid1()
        _Category = cboCategory.Text
        Panel1.Visible = False
        _PrintStatus = "Category"
    End Sub

    Private Sub ByCategoryToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ByCategoryToolStripMenuItem.Click
        Panel1.Visible = True
    End Sub

    Private Sub RefreshToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RefreshToolStripMenuItem.Click
        Call Load_Grid()
        _Category = ""
        _PrintStatus = "All"
        Panel1.Visible = False
        cboCategory.Text = ""

    End Sub
End Class