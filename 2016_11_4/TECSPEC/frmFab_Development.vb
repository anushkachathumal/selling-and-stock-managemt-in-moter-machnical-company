Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Public Class frmFab_Development
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim c_dataCustomer1 As DataTable
    Dim _CountryCode As String
    Dim _UnitCode As Integer
    Dim c_dataCustomer_Pr1 As System.Data.DataTable
    Dim c_dataCustomer_Pr2 As System.Data.DataTable
    Dim c_dataCustomer_Pr3 As System.Data.DataTable
    Dim c_dataCustomer4 As System.Data.DataTable
    Dim c_dataCustomer6 As System.Data.DataTable
    Dim c_dataCustomer7 As System.Data.DataTable
    Dim c_dataCustomer8 As System.Data.DataTable

    Function Load_Gride()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTableGrige_TECYARN
        UltraGrid1.DataSource = c_dataCustomer1
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 70
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
            .DisplayLayout.Bands(0).Columns(0).Width = 70
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
            .DisplayLayout.Bands(0).Columns(0).Width = 120
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 70
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False

        End With
    End Function


    Function Load_Gride_Boil_Result()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Dim I As Integer
        Dim vcWhere As String
        Dim _Code As Integer
        Dim T01 As DataSet
        Dim agroup1 As UltraGridGroup
        Dim agroup2 As UltraGridGroup
        Dim agroup3 As UltraGridGroup
        Dim agroup4 As UltraGridGroup
        Dim agroup5 As UltraGridGroup
        Dim agroup6 As UltraGridGroup
        Dim agroup7 As UltraGridGroup
        Dim agroup8 As UltraGridGroup
        Dim agroup9 As UltraGridGroup

        Dim _Date As Date
        Dim X As Integer
        Dim _coloumCount As Integer
        Dim Value As Double
        Dim _STSting As String
        Dim _week As Integer
        ' Dim i As Integer

        Try

            UltraGrid7.DisplayLayout.Bands(0).Groups.Clear()
            UltraGrid7.DisplayLayout.Bands(0).Columns.Dispose()

            'Dim agroup1 As UltraGridGroup
            'Dim agroup2 As UltraGridGroup
            'Dim agroup3 As UltraGridGroup
            'Dim agroup4 As UltraGridGroup
            'Dim agroup5 As UltraGridGroup
            '  Dim agroup6 As UltraGridGroup

            'If UltraGrid3.DisplayLayout.Bands(0).GroupHeadersVisible = True Then
            'Else
            '  agroup1.Key = ""
            '  agroup1 = UltraGrid3.DisplayLayout.Bands(0).Groups.Remove("GroupH")
            agroup1 = UltraGrid7.DisplayLayout.Bands(0).Groups.Add("Stage")
            '  agroup1 = dg_Knt_Pojection.DisplayLayout.Bands(0).Groups.te("GroupH")
            'Me.UltraGrid3.DisplayLayout.Bands(0).Columns.Add("Line", "Line Item")
            'Me.UltraGrid3.DisplayLayout.Bands(0).Columns("Line").Group = agroup1
            'Me.UltraGrid3.DisplayLayout.Bands(0).Columns("Line").Width = 50

            'Me.UltraGrid3.DisplayLayout.Bands(0).Columns.Add("##", "##")
            'Me.UltraGrid3.DisplayLayout.Bands(0).Columns("##").Group = agroup1
            'Me.UltraGrid3.DisplayLayout.Bands(0).Columns("##").Width = 120
            ''  End If
            ' agroup1 = UltraGrid3.DisplayLayout.Bands(0).Groups.Remove(0)


            agroup1.Width = 110
            Dim dt As DataTable = New DataTable()
            ' dt.Columns.Add("ID", GetType(Integer))
            Dim colWork As New DataColumn("##", GetType(String))
            dt.Columns.Add(colWork)
            colWork.ReadOnly = True

            For I = 1 To 6
                dt.Rows.Add("")
            Next

            'dt.Columns.Add("##", GetType(String))
            ' dt.Columns.Add("Shade", GetType(String))
           

            Me.UltraGrid7.SetDataBinding(dt, Nothing)
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns(0).Group = agroup1
            'Me.dg_YDP_Projection.DisplayLayout.Bands(0).Columns(1).Group = agroup1
            'Me.dg_YDP_Projection.DisplayLayout.Bands(0).Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns(0).Width = 180
            ' Me.dg_YDP_Projection.DisplayLayout.Bands(0).Columns(1).Width = 50
            Dim _Group As String
            'agroup2.Key = ""
            'agroup3.Key = ""
            'agroup4.Key = ""
            '' agroup5.Key = ""

         
            _Group = "Group_Weight"

            'agroup2.Key = ""
            agroup2 = UltraGrid7.DisplayLayout.Bands(0).Groups.Add("Group1")

            agroup2.Header.Caption = "Weight"
            agroup2.Width = 220
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns.Add("ProjectColumn", "Befor")
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns("ProjectColumn").Group = agroup2
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns("ProjectColumn").Width = 70
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns("ProjectColumn").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

            Me.UltraGrid7.DisplayLayout.Bands(0).Columns.Add("ProjectColumn1", "After")
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns("ProjectColumn1").Group = agroup2
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns("ProjectColumn1").Width = 70
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns("ProjectColumn1").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

            _Group = "Group_Width"

            'agroup2.Key = ""
            agroup3 = UltraGrid7.DisplayLayout.Bands(0).Groups.Add("Group2")

            agroup3.Header.Caption = "Width"
            agroup3.Width = 220
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns.Add("ProjectColumn_w", "Befor")
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns("ProjectColumn_w").Group = agroup3
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns("ProjectColumn_w").Width = 70
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns("ProjectColumn_w").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

            Me.UltraGrid7.DisplayLayout.Bands(0).Columns.Add("ProjectColumn_w1", "After")
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns("ProjectColumn_w1").Group = agroup3
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns("ProjectColumn_w1").Width = 70
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns("ProjectColumn_w1").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

            _Group = "Group_CPI"

            'agroup2.Key = ""
            agroup4 = UltraGrid7.DisplayLayout.Bands(0).Groups.Add("Group3")

            agroup4.Header.Caption = "CPI"
            agroup4.Width = 220
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns.Add("ProjectColumn_c", "Befor")
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns("ProjectColumn_c").Group = agroup4
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns("ProjectColumn_c").Width = 70
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns("ProjectColumn_c").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

            Me.UltraGrid7.DisplayLayout.Bands(0).Columns.Add("ProjectColumn_c1", "After")
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns("ProjectColumn_c1").Group = agroup4
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns("ProjectColumn_c1").Width = 70
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns("ProjectColumn_c1").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center


            _Group = "Group_wpi"

            'agroup2.Key = ""
            agroup5 = UltraGrid7.DisplayLayout.Bands(0).Groups.Add("Group4")
            agroup5.Header.Caption = "WPI"
            agroup5.Width = 220
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns.Add("ProjectColumn_p", "Befor")
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns("ProjectColumn_p").Group = agroup5
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns("ProjectColumn_p").Width = 70
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns("ProjectColumn_p").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

            Me.UltraGrid7.DisplayLayout.Bands(0).Columns.Add("ProjectColumn_p1", "After")
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns("ProjectColumn_p1").Group = agroup5
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns("ProjectColumn_p1").Width = 70
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns("ProjectColumn_p1").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center


            _Group = "Group_Test"

            'agroup2.Key = ""
            agroup6 = UltraGrid7.DisplayLayout.Bands(0).Groups.Add("Group5")
            agroup6.Header.Caption = "Test STD"
            agroup6.Width = 90

            Me.UltraGrid7.DisplayLayout.Bands(0).Columns.Add("ProjectColumn_t1", " ")
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns("ProjectColumn_t1").Group = agroup6
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns("ProjectColumn_t1").Width = 70
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns("ProjectColumn_t1").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

            _Group = "Group_E"

            'agroup2.Key = ""
            agroup7 = UltraGrid7.DisplayLayout.Bands(0).Groups.Add("Group6")
            agroup7.Header.Caption = "Elongation"
            agroup7.Width = 220
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns.Add("ProjectColumn_E", "Width")
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns("ProjectColumn_E").Group = agroup7
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns("ProjectColumn_E").Width = 70
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns("ProjectColumn_E").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

            Me.UltraGrid7.DisplayLayout.Bands(0).Columns.Add("ProjectColumn_E1", "Lenth")
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns("ProjectColumn_E1").Group = agroup7
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns("ProjectColumn_E1").Width = 70
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns("ProjectColumn_E1").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

            agroup8 = UltraGrid7.DisplayLayout.Bands(0).Groups.Add("Group7")
            agroup8.Header.Caption = "Moduler Width"
            agroup8.Width = 220
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns.Add("ProjectColumn_mW", "20%")
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns("ProjectColumn_MW").Group = agroup8
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns("ProjectColumn_MW").Width = 50
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns("ProjectColumn_MW").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

            Me.UltraGrid7.DisplayLayout.Bands(0).Columns.Add("ProjectColumn_mW1", "40%")
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns("ProjectColumn_MW1").Group = agroup8
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns("ProjectColumn_MW1").Width = 50
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns("ProjectColumn_MW1").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

            Me.UltraGrid7.DisplayLayout.Bands(0).Columns.Add("ProjectColumn_mW2", "60%")
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns("ProjectColumn_MW2").Group = agroup8
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns("ProjectColumn_MW2").Width = 50
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns("ProjectColumn_MW2").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

            Me.UltraGrid7.DisplayLayout.Bands(0).Columns.Add("ProjectColumn_mW3", "80%")
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns("ProjectColumn_MW3").Group = agroup8
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns("ProjectColumn_MW3").Width = 50
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns("ProjectColumn_MW3").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center


            agroup9 = UltraGrid7.DisplayLayout.Bands(0).Groups.Add("Group8")
            agroup9.Header.Caption = "Moduler Lenth"
            agroup9.Width = 220
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns.Add("ProjectColumn_mW4", "20%")
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns("ProjectColumn_MW4").Group = agroup9
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns("ProjectColumn_MW4").Width = 50
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns("ProjectColumn_MW4").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

            Me.UltraGrid7.DisplayLayout.Bands(0).Columns.Add("ProjectColumn_mW5", "40%")
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns("ProjectColumn_MW5").Group = agroup9
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns("ProjectColumn_MW5").Width = 50
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns("ProjectColumn_MW5").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

            Me.UltraGrid7.DisplayLayout.Bands(0).Columns.Add("ProjectColumn_mW6", "60%")
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns("ProjectColumn_MW6").Group = agroup9
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns("ProjectColumn_MW6").Width = 50
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns("ProjectColumn_MW6").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

            Me.UltraGrid7.DisplayLayout.Bands(0).Columns.Add("ProjectColumn_mW7", "80%")
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns("ProjectColumn_MW7").Group = agroup9
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns("ProjectColumn_MW7").Width = 50
            Me.UltraGrid7.DisplayLayout.Bands(0).Columns("ProjectColumn_MW7").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
                con.close()

            End If
        End Try
    End Function

    Function Load_Gride_Pr2()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer_Pr2 = CustomerDataClass.MakeDataTableGrige_Process
        dg_Pr2.DataSource = c_dataCustomer_Pr2
        With dg_Pr2
            .DisplayLayout.Bands(0).Columns(0).Width = 120
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 70
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False

        End With
    End Function

    Function Load_Gride_Pr3()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer_Pr3 = CustomerDataClass.MakeDataTableGrige_Process1
        UltraGrid2.DataSource = c_dataCustomer_Pr3
        With UltraGrid2
            .DisplayLayout.Bands(0).Columns(0).Width = 60
            ' .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 120
            '.DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
        End With
    End Function

    Function Load_Gride_Compation()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer7 = CustomerDataClass.MakeDataTableGrige_CompositionNew
        UltraGrid3.DataSource = c_dataCustomer7
        With UltraGrid3
            .DisplayLayout.Bands(0).Columns(0).Width = 120
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 70
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 70
            .DisplayLayout.Bands(0).Columns(3).Width = 70
            .DisplayLayout.Bands(0).Columns(4).Width = 70

        End With
    End Function

    Function Load_Gride_Compation1()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer8 = CustomerDataClass.MakeDataTableGrige_CompositionNew
        UltraGrid10.DataSource = c_dataCustomer8
        With UltraGrid10
            .DisplayLayout.Bands(0).Columns(0).Width = 120
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 70
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 70
            .DisplayLayout.Bands(0).Columns(3).Width = 70
            .DisplayLayout.Bands(0).Columns(4).Width = 70

        End With
    End Function

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

            '  UltraGrid4.DisplayLayout.Bands(0).Columns("Process Root").ValueList = Me.UltraDropDown3
            con.close()
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Function Load_Data_GridePr2()
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
                Dim newRow As DataRow = c_dataCustomer_Pr2.NewRow

                newRow("Process Root") = m01.Tables(0).Rows(i)("M51Dis")
                newRow("##") = False
                c_dataCustomer_Pr2.Rows.Add(newRow)

                i = i + 1
            Next

            '  UltraGrid4.DisplayLayout.Bands(0).Columns("Process Root").ValueList = Me.UltraDropDown3
            con.close()
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function
    Private Sub frmFab_Development_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_Gride()
        Call Load_Gride4()
        ' Call BindUltraDropDown()
        Call Load_Data_Gride4()
        ' Call BindUltraDropDown()
        'Call Load_Gride6()
        'Call Load_Data_Gride6()
        ' Call BindUltraDropDown_COMPOSION()

        Call Load_Gride_Pr1()
        Call Load_Gride_Pr2()
        Call Load_Gride_Pr3()

        Call Load_Data_GridePr2()
        Call Load_Gride_Compation()
        Call Load_Gride_Compation1()
        Call Load_Gride_Boil_Result()
    End Sub

    Private Sub UltraTabPageControl2_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles UltraTabPageControl2.Paint

    End Sub

    Private Sub UltraTabControl1_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles UltraTabControl1.SelectedTabChanged
        If UltraTabControl1.Tabs(0).Selected = True Then
            frmDisplay.Show()
        Else
            frmDisplay.Close()
        End If
    End Sub

  
    Private Sub UltraGrid4_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim I As Integer
        I = UltraGrid4.ActiveRow.Index
        'If Trim(lblA.Text) <> "" Then
        '    lblA.Text = Trim(lblA.Text) & "|" & Trim(UltraGrid4.Rows(I).Cells(0).Text)
        UltraGrid4.Rows(I).Cells(1).Value = True
        UltraGrid4.Rows(I).Cells(1).Appearance.BackColor = Color.Red
        UltraGrid4.Rows(I).Cells(0).Appearance.BackColor = Color.Red
        'Else
        'lblA.Text = Trim(UltraGrid4.Rows(I).Cells(0).Text)
        'UltraGrid4.Rows(I).Cells(1).Value = True
        'UltraGrid4.Rows(I).Cells(1).Appearance.BackColor = Color.Red
        'UltraGrid4.Rows(I).Cells(0).Appearance.BackColor = Color.Red
        'End If
    End Sub

    Private Sub UltraGrid4_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs)

    End Sub

    Private Sub UltraGrid4_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)

    End Sub

    Private Sub UltraButton9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton9.Click
        Me.Close()
    End Sub

    Private Sub dg_Pr2_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dg_Pr2.DoubleClick
        Dim I As Integer
        I = dg_Pr2.ActiveRow.Index
        ' If Trim(lblA1.Text) <> "" Then
        'lblA1.Text = Trim(lblA1.Text) & "|" & Trim(dg_Pr2.Rows(I).Cells(0).Text)
        dg_Pr2.Rows(I).Cells(1).Value = True
        dg_Pr2.Rows(I).Cells(1).Appearance.BackColor = Color.Red
        dg_Pr2.Rows(I).Cells(0).Appearance.BackColor = Color.Red
        'Else
        ''lblA1.Text = Trim(dg_Pr2.Rows(I).Cells(0).Text)
        'dg_Pr2.Rows(I).Cells(1).Value = True
        'dg_Pr2.Rows(I).Cells(1).Appearance.BackColor = Color.Red
        'dg_Pr2.Rows(I).Cells(0).Appearance.BackColor = Color.Red
        'End If
    End Sub

    Private Sub dg_Pr2_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles dg_Pr2.InitializeLayout

    End Sub

    Private Sub UltraGrid4_DoubleClick1(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim I As Integer
        Dim _Status As Boolean
        Dim _RootName As String

        _RootName = ""
        I = UltraGrid4.ActiveRow.Index
        ' If Trim(lblA1.Text) <> "" Then
        'lblA1.Text = Trim(lblA1.Text) & "|" & Trim(dg_Pr2.Rows(I).Cells(0).Text)
        UltraGrid4.Rows(I).Cells(1).Value = True
        UltraGrid4.Rows(I).Cells(1).Appearance.BackColor = Color.Red
        UltraGrid4.Rows(I).Cells(0).Appearance.BackColor = Color.Red
        UltraGrid4.Rows(I).Cells(1).Value = True
        _RootName = Trim(UltraGrid4.Rows(I).Cells(0).Text)
        _Status = False
        I = 0

        For Each uRow As UltraGridRow In UltraGrid2.Rows
            If _RootName = Trim(UltraGrid2.Rows(I).Cells(1).Text) Then
                _Status = True
                Exit For
            End If

            I = I + 1
        Next
        I = 0
        For Each uRow As UltraGridRow In UltraGrid2.Rows
            If Trim(UltraGrid2.Rows(I).Cells(1).Text) <> "" Then
            Else
                If _Status = False Then
                    UltraGrid2.Rows(I).Cells(1).Value = _RootName
                    Exit Sub
                End If
            End If
            I = I + 1
        Next
        I = UltraGrid2.Rows.Count
        I = I + 1
        If _Status = False Then
            Dim newRow As DataRow = c_dataCustomer_Pr3.NewRow
            newRow("Process No") = I
            newRow("##") = _RootName
            c_dataCustomer_Pr3.Rows.Add(newRow)
        End If
    End Sub

    Private Sub UltraGrid4_InitializeLayout_1(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs)
        e.Layout.AddNewBox.Hidden = True
    End Sub

    Private Sub UltraGrid2_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs)

    End Sub

    Private Sub UltraTabPageControl1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles UltraTabPageControl1.Paint

    End Sub

    Private Sub UltraGrid2_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        On Error Resume Next
        Dim A As String
        Dim i As Integer
        Dim X As Integer
        If e.KeyCode = Keys.F1 Then
            i = UltraGrid2.ActiveRow.Index
            Dim newRow As DataRow = c_dataCustomer_Pr3.NewRow
            newRow("Process No") = UltraGrid2.Rows.Count + 1
            newRow("##") = ""
            c_dataCustomer_Pr3.Rows.Add(newRow)

            X = UltraGrid2.Rows.Count
            For Each uRow As UltraGridRow In UltraGrid2.Rows
                If X - 1 = i Then
                    'UltraGrid2.Rows(X - 1).Cells(0).Value = X
                    UltraGrid2.Rows(X - 1).Cells(1).Value = ""
                    Exit For
                Else

                    'UltraGrid2.Rows(X - 1).Cells(0).Value = X
                    ' MsgBox(UltraGrid2.Rows(X - 2).Cells(1).Text)
                    UltraGrid2.Rows(X - 1).Cells(1).Value = UltraGrid2.Rows(X - 2).Cells(1).Text
                End If
                X = X - 1
            Next
        End If
    End Sub

    Private Sub Panel21_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel21.Paint

    End Sub

    Private Sub UltraButton18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        With UltraTabControl2
            .Tabs(7).Visible = True
            UltraTabControl2.SelectedTab = UltraTabControl2.Tabs(7)
        End With
    End Sub
End Class