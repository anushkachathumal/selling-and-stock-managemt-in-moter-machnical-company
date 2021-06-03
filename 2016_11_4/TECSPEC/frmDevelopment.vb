Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Public Class frmDevelopment
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim c_dataCustomer1 As DataTable
    Dim _CountryCode As String
    Dim _UnitCode As Integer
    Dim c_dataCustomer_Pr1 As System.Data.DataTable
    Dim c_dataCustomer_Pr2 As System.Data.DataTable
    Dim c_dataCustomer4 As System.Data.DataTable
    Dim c_dataCustomer6 As System.Data.DataTable

    Function Load_Gride()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTableGrige_TECYARN
        UltraGrid1.DataSource = c_dataCustomer1
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 270
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 70
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 60
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(3).Width = 80
            .DisplayLayout.Bands(0).Columns(3).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(4).Width = 80
            .DisplayLayout.Bands(0).Columns(4).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(5).Width = 80
            .DisplayLayout.Bands(0).Columns(5).AutoEdit = True
            '   .DisplayLayout.Bands(0).Columns(6).AutoEdit = False
            ' .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Left
            .DisplayLayout.Bands(0).Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            '.DisplayLayout.Bands(0).Columns(3).Width = 90
            '.DisplayLayout.Bands(0).Columns(4).Width = 90
            '.DisplayLayout.Bands(0).Columns(5).Width = 90
            ''  .DisplayLayout.Bands(0).Columns(6).Width = 90
            ' .DisplayLayout.Bands(0).Columns(7).Width = 90

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function

    Function Load_Gride_Pr1()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer_Pr1 = CustomerDataClass.MakeDataTableGrige_TECYARN
        dg_Pr1.DataSource = c_dataCustomer_Pr1
        With dg_Pr1
            .DisplayLayout.Bands(0).Columns(0).Width = 270
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 70
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 60
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(3).Width = 80
            .DisplayLayout.Bands(0).Columns(3).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(4).Width = 80
            .DisplayLayout.Bands(0).Columns(4).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(5).Width = 80
            .DisplayLayout.Bands(0).Columns(5).AutoEdit = True
            '   .DisplayLayout.Bands(0).Columns(6).AutoEdit = False
            ' .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Left
            .DisplayLayout.Bands(0).Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            '.DisplayLayout.Bands(0).Columns(3).Width = 90
            '.DisplayLayout.Bands(0).Columns(4).Width = 90
            '.DisplayLayout.Bands(0).Columns(5).Width = 90
            ''  .DisplayLayout.Bands(0).Columns(6).Width = 90
            ' .DisplayLayout.Bands(0).Columns(7).Width = 90

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function

    Function Load_Gride4()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer4 = CustomerDataClass.MakeDataTableGrige_Process
        UltraGrid4.DataSource = c_dataCustomer4
        With UltraGrid4
            .DisplayLayout.Bands(0).Columns(0).Width = 170
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 70
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
           
        End With
    End Function

    Function Load_Gride_Pr2()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer_Pr2 = CustomerDataClass.MakeDataTableGrige_Process
        dg_Pr2.DataSource = c_dataCustomer_Pr2
        With dg_Pr2
            .DisplayLayout.Bands(0).Columns(0).Width = 170
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 70
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False

        End With
    End Function

    Function Load_Gride6()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer6 = CustomerDataClass.MakeDataTableGrige_Composition
        UltraGrid6.DataSource = c_dataCustomer6
        With UltraGrid6
            .DisplayLayout.Bands(0).Columns(0).Width = 130
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 70
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False

        End With
    End Function

    Private Sub frmDevelopment_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_Gride()
        Call Load_Gride4()
        Call BindUltraDropDown()
        Call Load_Data_Gride4()
        Call BindUltraDropDown()
        Call Load_Gride6()
        Call Load_Data_Gride6()
        Call BindUltraDropDown_COMPOSION()

        Call Load_Gride_Pr1()
        Call Load_Gride_Pr2()

    End Sub

    Function Load_Data_Gride4()
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim m01 As DataSet
        Dim Sql As String
        Dim i As Integer

        Try
            Sql = "select * from M51Process_Root order by M51Code"
            m01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            For Each DTRow4 As DataRow In m01.Tables(0).Rows
                'tmpQty = 0
                Dim newRow As DataRow = c_dataCustomer4.NewRow

                newRow("Process Root") = m01.Tables(0).Rows(i)("M51Dis")
                newRow("##") = False
                c_dataCustomer4.Rows.Add(newRow)

                i = i + 1
            Next
          
            UltraGrid4.DisplayLayout.Bands(0).Columns("Process Root").ValueList = Me.UltraDropDown3
            con.close()
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Function Load_Data_Gride6()
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim m01 As DataSet
        Dim Sql As String
        Dim i As Integer

        Try
           
            For i = 1 To 4
                'tmpQty = 0
                Dim newRow As DataRow = c_dataCustomer6.NewRow

                newRow("Type") = ""
                ' newRow("##") = False
                c_dataCustomer6.Rows.Add(newRow)

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

    Private Sub BindUltraDropDown()
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim m01 As DataSet
        Dim Sql As String
        Dim i As Integer
        Dim dt As DataTable = New DataTable()
        Try
            Sql = "select * from M51Process_Root order by M51Code"
            m01 = DBEngin.ExecuteDataset(con, Nothing, Sql)

            ' dt.Columns.Add("ID", GetType(Integer))
            dt.Columns.Add("##", GetType(String))
            For Each DTRow4 As DataRow In m01.Tables(0).Rows
                dt.Rows.Add(New Object() {Trim(m01.Tables(0).Rows(i)("M51Dis"))})
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

    Private Sub BindUltraDropDown_COMPOSION()
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim m01 As DataSet
        Dim Sql As String
        Dim i As Integer
        Dim dt As DataTable = New DataTable()
        Try
            Sql = "select * from M53Type order by M53Code"
            m01 = DBEngin.ExecuteDataset(con, Nothing, Sql)

            ' dt.Columns.Add("ID", GetType(Integer))
            dt.Columns.Add("##", GetType(String))
            For Each DTRow4 As DataRow In m01.Tables(0).Rows
                dt.Rows.Add(New Object() {Trim(m01.Tables(0).Rows(i)("M53Dis"))})
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

    Private Sub UltraGrid4_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles UltraGrid4.InitializeLayout
        e.Layout.Bands(0).Columns("Process Root").ValueList = Me.UltraDropDown3
    End Sub

    Private Sub UltraTabControl1_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles UltraTabControl1.SelectedTabChanged
        If UltraTabControl1.Tabs(1).Selected = True Then
            frmDisplay.Show()
        Else
            frmDisplay.Close()
        End If
    End Sub

    Private Sub UltraTabPageControl1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles UltraTabPageControl1.Paint

    End Sub

    Private Sub UltraGrid6_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles UltraGrid6.InitializeLayout
        e.Layout.Bands(0).Columns("Type").ValueList = Me.UltraDropDown1
    End Sub
End Class