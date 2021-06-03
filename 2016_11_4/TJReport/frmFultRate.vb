
Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.DAL_frmWinner
Imports DBLotVbnet.common
Imports DBLotVbnet.MDIMain
Imports System.Net.NetworkInformation
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.Quarrys
Imports System.IO.File
Imports System.IO.StreamWriter
Imports System.Net.Mail
Imports Microsoft.Office.Interop.Excel
Imports System.Drawing
Imports Spire.XlS
Public Class frmFultRate
    Dim oFile As System.IO.File
    Dim oWrite As System.IO.StreamWriter
    Dim exc As New Application
    Dim Clicked As String
    Dim workbooks As Workbooks = exc.Workbooks
    Dim workbook As _Workbook = Workbooks.Add(XlWBATemplate.xlWBATWorksheet)
    Dim sheets As Sheets = Workbook.Worksheets
    Dim worksheet1 As _Worksheet = CType(sheets.Item(1), _Worksheet)


    Private Sub frmFultRate_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtDate.Text = Today
        txtTo.Text = Today

        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select M03Quality as [Quality] from M03Knittingorder group by M03Quality order by M03Quality "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboQuality
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 190

            End With
            '-----------------------------------------------------------------------------------------------------
            'Yarn Supplier

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try

    End Sub

    Private Sub cboFrate1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboFrate1.CheckedChanged
        If cboFrate1.Checked = True Then
            cboFrate2.Checked = False
            cboFrate3.Checked = False
            cboFrate4.Checked = False
        End If
    End Sub

    Private Sub cboFrate2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboFrate2.CheckedChanged
        If cboFrate2.Checked = True Then
            cboFrate1.Checked = False
            cboFrate3.Checked = False
            cboFrate4.Checked = False
        End If
    End Sub

    Private Sub cboFrate3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboFrate3.CheckedChanged
        If cboFrate3.Checked = True Then
            cboFrate2.Checked = False
            cboFrate1.Checked = False
            cboFrate4.Checked = False
        End If
    End Sub

    Private Sub cboFrate4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboFrate4.CheckedChanged
        If cboFrate4.Checked = True Then
            cboFrate2.Checked = False
            cboFrate3.Checked = False
            cboFrate1.Checked = False
        End If
    End Sub

    Private Sub cmdEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEdit.Click
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet

        If chk0.Checked = True And cboQuality.Text <> "" Then

        End If
    End Sub
End Class