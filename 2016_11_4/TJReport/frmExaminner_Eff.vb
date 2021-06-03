
Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports System.Drawing.Color
Imports CrystalDecisions.CrystalReports.Engine
Imports System.Configuration
Imports Microsoft.Office.Interop.Excel
Imports System.Drawing
Imports Spire.XlS

Public Class frmExaminner_Eff
    Dim Clicked As String
    Private Sub frmExaminner_Eff_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtYear.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
    End Sub

    Function Load_Month()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Try
            Sql = "select M13Name as [Month] from M13Month"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)

            With cboMonth
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 110
                .Rows.Band.Columns(1).Width = 160
          

            End With
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""


        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function


    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        Clicked = "ADD"
        OPR2.Enabled = True
        ' cboDep.ToggleDropdown()
        cmdEdit.Enabled = True
    End Sub
    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub
    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        common.ClearAll(OPR2)
        Clicked = ""
        cmdAdd.Enabled = True
        'cmdSave.Enabled = False
        cmdAdd.Focus()
    End Sub


    Private Sub cmdEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEdit.Click

    End Sub
End Class