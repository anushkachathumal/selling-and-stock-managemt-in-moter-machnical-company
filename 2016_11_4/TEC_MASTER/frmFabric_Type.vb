Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Public Class frmFabric_Type
    Dim c_dataCustomer1 As DataTable

    Private Sub UltraButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton3.Click
        Me.Close()
    End Sub

    Function Load_Gride()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTableGrige_Fabric
        UltraGrid1.DataSource = c_dataCustomer1
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 70
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 240
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False

        End With
    End Function

    Private Sub frmFabric_Type_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_Gride()
    End Sub
End Class