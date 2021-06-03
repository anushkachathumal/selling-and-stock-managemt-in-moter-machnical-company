Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Public Class frmDisplay2
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim c_dataCustomer1 As DataTable

    Function Load_Gride()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTableGrige_Screen2
        UltraGrid4.DataSource = c_dataCustomer1
        With UltraGrid4
            .DisplayLayout.Bands(0).Columns(0).Width = 80
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 80
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 80
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
          
        End With
    End Function

    Private Sub frmDisplay2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_Gride()
        Call Load_Data_Gride6()
        Call BindUltraDropDown()
        Call BindUltraDropDown_2()
    End Sub

    Private Sub BindUltraDropDown()
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim m01 As DataSet
        Dim Sql As String
        Dim i As Integer
        Dim dt As DataTable = New DataTable()
        Try
            Sql = "select * from M54Special WHERE M54Status='1' order by M54Code"
            m01 = DBEngin.ExecuteDataset(con, Nothing, Sql)

            ' dt.Columns.Add("ID", GetType(Integer))
            dt.Columns.Add("##", GetType(String))
            For Each DTRow4 As DataRow In m01.Tables(0).Rows
                dt.Rows.Add(New Object() {Trim(m01.Tables(0).Rows(i)("M54Des"))})
                i = i + 1
            Next
            dt.AcceptChanges()

            Me.UltraDropDown3.SetDataBinding(dt, Nothing)
            '  Me.UltraDropDown1.ValueMember = "ID"
            Me.UltraDropDown3.DisplayMember = "##"
            UltraDropDown3.Rows.Band.Columns(0).Width = 150
            con.close()
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Sub

    Function Load_Data_Gride6()
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim m01 As DataSet
        Dim Sql As String
        Dim i As Integer

        Try

            For i = 1 To 16
                'tmpQty = 0
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                newRow("SPECIAL") = ""
                ' newRow("##") = False
                c_dataCustomer1.Rows.Add(newRow)

                'i = i + 1
            Next

            ' UltraGrid4.DisplayLayout.Bands(0).Columns("Process Root").ValueList = Me.UltraDropDown3
            con.close()
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Private Sub UltraGrid4_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles UltraGrid4.InitializeLayout
        e.Layout.Bands(0).Columns("SPECIAL").ValueList = Me.UltraDropDown3
        e.Layout.Bands(0).Columns("STRUCTURE").ValueList = Me.UltraDropDown1
    End Sub


    Private Sub BindUltraDropDown_2()
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim m01 As DataSet
        Dim Sql As String
        Dim i As Integer
        Dim dt As DataTable = New DataTable()
        Try
            Sql = "select * from M54Special WHERE M54Status='2' order by M54Code"
            m01 = DBEngin.ExecuteDataset(con, Nothing, Sql)

            ' dt.Columns.Add("ID", GetType(Integer))
            dt.Columns.Add("##", GetType(String))
            For Each DTRow4 As DataRow In m01.Tables(0).Rows
                dt.Rows.Add(New Object() {Trim(m01.Tables(0).Rows(i)("M54Des"))})
                i = i + 1
            Next
            dt.AcceptChanges()

            Me.UltraDropDown1.SetDataBinding(dt, Nothing)
            '  Me.UltraDropDown1.ValueMember = "ID"
            Me.UltraDropDown1.DisplayMember = "##"
            UltraDropDown1.Rows.Band.Columns(0).Width = 150
            con.close()
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Sub
End Class