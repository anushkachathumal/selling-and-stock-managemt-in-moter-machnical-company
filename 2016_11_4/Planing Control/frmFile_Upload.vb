
Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports System.IO
Public Class frmFile_Upload
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim c_dataCustomer1 As System.Data.DataTable
    Dim c_dataCustomer2 As System.Data.DataTable

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    'Function Load_Gride()
    '    Dim CustomerDataClass As New DAL_InterLocation()
    '    c_dataCustomer1 = MakeDataTable_StockCode()
    '    UltraGrid1.DataSource = c_dataCustomer1
    '    With UltraGrid1
    '        .DisplayLayout.Bands(0).Columns(0).Width = 90
    '        .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
    '        .DisplayLayout.Bands(0).Columns(1).Width = 120
    '        .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
    '        .DisplayLayout.Bands(0).Columns(2).Width = 90
    '        '.DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

    '        '.DisplayLayout.Bands(0).Columns(3).Width = 60
    '        '.DisplayLayout.Bands(0).Columns(5).Width = 60
    '        '.DisplayLayout.Bands(0).Columns(8).Width = 60
    '        '.DisplayLayout.Bands(0).Columns(7).Width = 70
    '        '.DisplayLayout.Bands(0).Columns(9).Width = 60

    '    End With
    'End Function
End Class