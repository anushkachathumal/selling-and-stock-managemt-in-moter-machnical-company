Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports Excel = Microsoft.Office.Interop.Excel
'Imports Office = Microsoft.Office.Core
Imports Microsoft.Office.Interop.Outlook
Imports System.Drawing
Imports Spire.XlS
Imports System.Xml
Imports Infragistics.Win.UltraWinGrid.RowLayoutStyle.GroupLayout
Imports Infragistics.Win.UltraWinToolTip
Imports Infragistics.Win.FormattedLinkLabel
Imports Infragistics.Win.Misc
Imports System.Diagnostics
Public Class frmtjl_Merchant
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim c_dataCustomer1 As DataTable
    Dim _EPF As String
    Dim _Email As String
    Dim _LeadTime As String

    Dim c_dataCustomer As DataTable
    Dim c_dataCustomer2 As DataTable
    Function Load_Gride_SalesOrder()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer = CustomerDataClass.MakeDataTable_Sales_Order
        UltraGrid1.DataSource = c_dataCustomer
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(1).Width = 50
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(3).Width = 190
            .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(3).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(3).Width = 60
            .DisplayLayout.Bands(0).Columns(4).Width = 60
            .DisplayLayout.Bands(0).Columns(6).Width = 60
            .DisplayLayout.Bands(0).Columns(9).Width = 60
            .DisplayLayout.Bands(0).Columns(8).Width = 70
            .DisplayLayout.Bands(0).Columns(11).Width = 60
            .DisplayLayout.Bands(0).Columns(18).Width = 60
            .DisplayLayout.Bands(0).Columns(19).Width = 130
            .DisplayLayout.Bands(0).Columns(20).Width = 90
        End With
    End Function

    Function Load_Sales_Order()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        'Load sales order to cboSO combobox

        Try
            Sql = "select CONVERT(INT,M01Sales_Order) as [Sales Order],max(M01Cuatomer_Name) as [Customer] from M01Sales_Order_SAP WHERE M01SO_Qty<>M01Delivary_Qty and M01Status='A' group by M01Sales_Order "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboSO
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 160
                .Rows.Band.Columns(1).Width = 260


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
    Function Load_PLANNER()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        'Load Planer to cboPlanner combobox

        Try
            Sql = "select Name as [Planner's Name] from users where Designation='PLANNER' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboPlaner
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 160
                '.Rows.Band.Columns(1).Width = 260


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

    
    Private Sub frmtjl_Merchant_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_PLANNER()
        Call Load_Sales_Order()
        'Call Load_Gride_SalesOrder()
        Call Load_Grid_Main()
    End Sub

    Function Load_Grid_Main()
        Dim agroup1 As UltraGridGroup
        Dim agroup2 As UltraGridGroup
        Dim agroup3 As UltraGridGroup
        Dim agroup4 As UltraGridGroup
        Dim agroup5 As UltraGridGroup
        Dim agroup6 As UltraGridGroup
        Dim agroup7 As UltraGridGroup
        Dim agroup8 As UltraGridGroup
        Dim agroup9 As UltraGridGroup

        agroup1 = UltraGrid1.DisplayLayout.Bands(0).Groups.Add("##")
        agroup1.Width = 110
        Dim dt As DataTable = New DataTable()
        ' dt.Columns.Add("ID", GetType(Integer))
        Dim colWork As New DataColumn("##", GetType(String))
        dt.Columns.Add(colWork)
        colWork.ReadOnly = True
        colWork = New DataColumn("Line Item", GetType(String))
        colWork.MaxLength = 80
        dt.Columns.Add(colWork)
        colWork.ReadOnly = True
        colWork = New DataColumn("Matching", GetType(String))
        colWork.MaxLength = 80
        dt.Columns.Add(colWork)
        colWork = New DataColumn("Quality", GetType(String))
        colWork.MaxLength = 80
        dt.Columns.Add(colWork)
        colWork.ReadOnly = True
        colWork = New DataColumn("Retailer", GetType(String))
        colWork.MaxLength = 80
        dt.Columns.Add(colWork)
        colWork.ReadOnly = True
        colWork = New DataColumn("FG Stock", GetType(String))
        colWork.MaxLength = 80
        dt.Columns.Add(colWork)
        colWork.ReadOnly = True
        colWork = New DataColumn("Greige Qty(kg)", GetType(String))
        colWork.MaxLength = 80
        dt.Columns.Add(colWork)
        colWork = New DataColumn("Qty(Kg)", GetType(String))
        colWork.MaxLength = 80
        dt.Columns.Add(colWork)
        colWork.ReadOnly = True
        colWork = New DataColumn("Greige Booking", GetType(String))
        colWork.MaxLength = 80
        dt.Columns.Add(colWork)
        colWork = New DataColumn("Greige Booking No", GetType(String))
        colWork.MaxLength = 80
        dt.Columns.Add(colWork)

        colWork = New DataColumn("Yarn Booking", GetType(String))
        colWork.MaxLength = 80
        dt.Columns.Add(colWork)
        colWork = New DataColumn("Yarn Booking No", GetType(String))
        colWork.MaxLength = 80
        dt.Columns.Add(colWork)
        colWork = New DataColumn("1st Bulk", GetType(String))
        colWork.MaxLength = 80
        dt.Columns.Add(colWork)
        colWork = New DataColumn("Submission", GetType(String))
        colWork.MaxLength = 80
        dt.Columns.Add(colWork)

        Me.UltraGrid1.SetDataBinding(dt, Nothing)
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns(0).Group = agroup1
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns(1).Group = agroup1
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns(2).Group = agroup1
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns(3).Group = agroup1
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns(4).Group = agroup1
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns(5).Group = agroup1
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns(6).Group = agroup1
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns(7).Group = agroup1
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns(8).Group = agroup1
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns(9).Group = agroup1
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns(10).Group = agroup1
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns(11).Group = agroup1
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns(12).Group = agroup1
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns(13).Group = agroup1
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
        '  Me.UltraGrid1.DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns(0).Width = 40
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns(1).Width = 60
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns(2).Width = 80
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns(3).Width = 80
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns(4).Width = 150
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns(5).Width = 90
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns(6).Width = 90
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns(7).Width = 90
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns(8).Width = 90
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns(9).Width = 110
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns(10).Width = 90
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns(11).Width = 110
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns(12).Width = 90
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns(13).Width = 110

        agroup2 = UltraGrid1.DisplayLayout.Bands(0).Groups.Add("NPL")
        agroup2.Header.Caption = "NPL"
        agroup2.Width = 220
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns.Add("NA", "Approved")
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns("NA").Group = agroup2
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns("NA").Width = 70
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns("NA").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Left
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns.Add("ND", "Date")
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns("ND").Group = agroup2
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns("ND").Width = 70
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns("ND").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Left

        agroup3 = UltraGrid1.DisplayLayout.Bands(0).Groups.Add("PP")
        agroup3.Header.Caption = "PP"
        agroup3.Width = 220
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns.Add("PA", "Approved")
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns("PA").Group = agroup3
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns("PA").Width = 70
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns("PA").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Left
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns.Add("PD", "Date")
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns("PD").Group = agroup3
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns("PD").Width = 70
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns("PD").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Left

        agroup4 = UltraGrid1.DisplayLayout.Bands(0).Groups.Add("LD")
        agroup4.Header.Caption = "Lab dip"
        agroup4.Width = 220
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns.Add("LA", "Approved")
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns("LA").Group = agroup4
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns("LA").Width = 70
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns("LA").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Left
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns.Add("LD", "Date")
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns("LD").Group = agroup4
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns("LD").Width = 70
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns("LD").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Left

        agroup5 = UltraGrid1.DisplayLayout.Bands(0).Groups.Add("CL")
        agroup5.Header.Caption = "Coloring Order"
        agroup5.Width = 220
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns.Add("CA", " ")
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns("CA").Group = agroup5
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns("CA").Width = 70
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns("CA").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Left

        agroup6 = UltraGrid1.DisplayLayout.Bands(0).Groups.Add("ST")
        agroup6.Header.Caption = "Strategic Order"
        agroup6.Width = 220
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns.Add("SA", "Available")
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns("SA").Group = agroup6
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns("SA").Width = 70
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns("SA").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Left
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns.Add("SO", "Order No")
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns("SO").Group = agroup6
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns("SO").Width = 120
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns("SO").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Left
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns.Add("SR", "Rolling")
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns("SR").Group = agroup6
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns("SR").Width = 70
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns("SR").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Left
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns.Add("SN", "Non Rolling")
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns("SN").Group = agroup6
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns("SN").Width = 70
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns("SN").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Left

        agroup7 = UltraGrid1.DisplayLayout.Bands(0).Groups.Add("IQ")
        agroup7.Header.Caption = "Inquiry Order"
        agroup7.Width = 220
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns.Add("IA", "Available")
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns("IA").Group = agroup7
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns("IA").Width = 70
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns("IA").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Left
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns.Add("IO", "Order No")
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns("IO").Group = agroup7
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns("IO").Width = 120
        Me.UltraGrid1.DisplayLayout.Bands(0).Columns("IO").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Left
    End Function
    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Private Sub UltraGrid1_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles UltraGrid1.InitializeLayout

    End Sub
End Class