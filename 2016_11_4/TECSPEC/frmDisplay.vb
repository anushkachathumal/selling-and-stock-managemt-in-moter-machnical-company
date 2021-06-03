Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Public Class frmDisplay

    Private Sub UltraButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton1.Click
        Me.Close()
    End Sub

    Function Load_Combo()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        'Load sales order to cboSO combobox

        Try
            Sql = "select M52Dis as [##] from M52Spec order by M52Code "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboDis
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 270
                ' .Rows.Band.Columns(1).Width = 260


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

    Private Sub frmDisplay_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_Combo()
    End Sub

    Function Search_Combo()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        'Load sales order to cboSO combobox

        Try
            Sql = "select * from M52Spec where M52Dis='" & cboDis.Text & "' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                If Trim(M01.Tables(0).Rows(0)("M52Code")) = "1" Then
                    lblDis.Visible = True
                    cboQuality.Visible = True
                    lblDis.Text = "T Series"
                ElseIf Trim(M01.Tables(0).Rows(0)("M52Code")) = "2" Then

                    lblDis.Visible = True
                    cboQuality.Visible = True
                    lblDis.Text = "Quality No"
                ElseIf Trim(M01.Tables(0).Rows(0)("M52Code")) = "3" Then

                    lblDis.Visible = True
                    cboQuality.Visible = True
                    lblDis.Text = "Quality No"
                ElseIf Trim(M01.Tables(0).Rows(0)("M52Code")) = "4" Then

                    lblDis.Visible = True
                    cboQuality.Visible = True
                    lblDis.Text = "Quality No"
                End If
            End If
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

    Private Sub cboDis_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles cboDis.InitializeLayout

    End Sub

    Private Sub cboDis_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboDis.TextChanged
        Call Search_Combo()
    End Sub

    Private Sub cmdDis_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDis.Click
        Me.Close()
        frmQuality_Creation.Show()
    End Sub
End Class