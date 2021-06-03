Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Public Class frmKNT_Spec
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim c_dataCustomer1 As DataTable
    Dim _CountryCode As String
    Dim _UnitCode As Integer
    Dim c_dataCustomer_Pr1 As System.Data.DataTable
    Dim c_dataCustomer_Pr2 As System.Data.DataTable
    Dim c_dataCustomer_Pr3 As System.Data.DataTable
    Private Sub frmKNT_Spec_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_Gride_Compation1()
    End Sub
    Private Sub Panel10_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel10.Paint

    End Sub

    Private Sub UltraGroupBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraGroupBox1.Click

    End Sub
    Function Load_Gride_Compation1()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer_Pr1 = CustomerDataClass.MakeDataTableGrige_KNTSpec_Yarn
        UltraGrid3.DataSource = c_dataCustomer_Pr1
        With UltraGrid3
            .DisplayLayout.Bands(0).Columns(0).Width = 120
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 80
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 70
            .DisplayLayout.Bands(0).Columns(3).Width = 70
            .DisplayLayout.Bands(0).Columns(4).Width = 70

        End With
    End Function
End Class