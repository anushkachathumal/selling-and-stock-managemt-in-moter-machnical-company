Imports System.Drawing
Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation

Public Class frmKnt_Plan
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim c_dataCustomer1 As System.Data.DataTable

    Private Sub frmKnt_Plan_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        frmGriege_Stock.chkKnt_Plan.Checked = False
    End Sub

    Private Sub frmKnt_Plan_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim X As Integer
        Dim y As Integer

        X = Me.Width
        grp1.Width = X - 50
        Panel1.Width = X - 50
        Panel1.Height = 200
        grp1.Height = 200

        Call Create_MCGroup()


    End Sub

  
    Function Create_MCGroup()
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
        Dim _SizeX As Integer
        Dim _SizeY As Integer

        _SizeX = 20
        _SizeY = 45
        i = 0

        M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetKnitting_PLN", New SqlParameter("@cQryType", "MCG"))
        For Each DTRow3 As DataRow In M01.Tables(0).Rows
            Dim B As New Infragistics.Win.UltraWinEditors.UltraTextEditor
            grp1.Controls.Add(B)
            B.Width = 90
            B.Height = 60
            B.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            B.Text = M01.Tables(0).Rows(i)("M32GroupName")
            B.ReadOnly = True
            B.Location = New Point(_SizeX, _SizeY)


            _SizeY = _SizeY + 40

            i = i + 1
        Next

        grp1.Height = _SizeY + 100
        Panel1.Height = _SizeY + 100
    End Function
   
    Private Sub grp1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles grp1.Click

    End Sub
End Class