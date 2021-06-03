Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports System.Drawing.Color
Imports CrystalDecisions.CrystalReports.Engine
Imports System.Configuration
Public Class frmProduction_Examinner
    Dim Clicked As String
    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        Clicked = "ADD"
        OPR0.Enabled = True
        ' cboDep.ToggleDropdown()

        txtDate.Text = Today
        txtTo.Text = Today
        cmdEdit.Enabled = True
    End Sub
    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        common.ClearAll(OPR0)
        Clicked = ""
        cmdAdd.Enabled = True
        'cmdSave.Enabled = False
        cmdAdd.Focus()
    End Sub

    Private Sub frmProduction_Examinner_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim recArea As DataSet
        Dim M01 As DataSet

        Try
            'SET COMPANY
            Sql = "select EPFNo as [EPF No] from Users "
            recArea = DBEngin.ExecuteDataset(con, Nothing, Sql)
            cboCategory.DataSource = recArea
            cboCategory.Rows.Band.Columns(0).Width = 370
            ' cboSupp.Rows.Band.Columns(1).Width = 170



            txtDate.Text = Today
            txtTo.Text = Today

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Sub

    Private Sub cmdEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEdit.Click
        Dim B As New ReportDocument
        Dim A As String
        Dim StrFromDate As String
        Dim StrToDate As String

        Try
            StrFromDate = "(" & Year(txtDate.Value) & ", " & VB6.Format(Month(txtDate.Value), "0#") & ", " & VB6.Format(CDate(txtDate.Text).Day, "0#") & ", 00, 00, 00)"
            StrToDate = "(" & Year(txtTo.Value) & ", " & VB6.Format(Month(txtTo.Value), "0#") & ", " & VB6.Format(CDate(txtTo.Text).Day, "0#") & ", 00, 00, 00)"


            A = ConfigurationManager.AppSettings("ReportPath") + "\rptProEPF.rpt"
            B.Load(A.ToString)
            B.SetDatabaseLogon("sa", "tommya")
            B.SetParameterValue("To", txtTo.Text)
            B.SetParameterValue("From", txtDate.Text)
            B.SetParameterValue("Emp", txtM2.Text)

            frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
            frmReport.CrystalReportViewer1.DisplayToolbar = True
            ' frmReport.CrystalReportViewer1.SelectionFormula = "{R01Report.R01WorkStation}='" & netCard & "'"
            
            frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01InsEPF}='" & Trim(cboCategory.Text) & "' and {T01Transaction_Header.T01Date} in DateTime " & StrFromDate & " to DateTime " & StrToDate & ""
            frmReport.Refresh()
            ' frmReport.CrystalReportViewer1.PrintReport()
            ' B.PrintToPrinter(1, True, 0, 0)
            frmReport.MdiParent = MDIMain
            frmReport.Show()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Sub

    Function Search_Employeename()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim recArea As DataSet
        Dim M01 As DataSet

        Try
            'SET COMPANY
            Sql = "select * from Users where EPFNo='" & cboCategory.Text & "'"
            recArea = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(recArea) Then
                txtM2.Text = recArea.Tables(0).Rows(0)("FirstName")
            End If

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub cboCategory_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboCategory.AfterCloseUp
        Call Search_Employeename()
    End Sub

    Private Sub cboCategory_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles cboCategory.InitializeLayout

    End Sub

    Private Sub cboCategory_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboCategory.KeyUp
        If e.KeyCode = 13 Then
            Call Search_Employeename()
            cmdEdit.Focus()

        End If
    End Sub

    Private Sub OPR0_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OPR0.Click

    End Sub
End Class