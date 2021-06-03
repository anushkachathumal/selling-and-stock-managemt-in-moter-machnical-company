Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports System.IO.StreamReader
Imports DBLotVbnet.delivary_forcast
'Imports Microsoft.Office.Interop.Excel

Public Class frmDelivary_Forcust


    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim c_dataCustomer1 As System.Data.DataTable
    Dim _Customer As String
    Dim _Department As String
    Dim _Merchant As String
    Dim c_dataCustomer As DataTable
    Dim _ProDis1 As String
    Dim _ProDis2 As String
    Dim _ProDis3 As String
    Dim _ProQ1 As Double
    Dim _ProQ2 As Double
    Dim _ProQ3 As Double
    Dim _StockDis1 As String
    Dim _StockDis2 As String
    Dim _StockDis3 As String
    Dim _Promiss_To_delTot As Double

    Dim _StockWK1 As String
    Dim _StockWK2 As String

    Dim _TobeDis As String
    Dim _TobeDis1 As String
    Dim _TobeWkQty As Double
    Dim _TobeWKQty1 As Double
    Dim _WIPDye As Double
    Dim _WIPExam As Double
    Dim _WIPFinishing As Double
    Const MAX_SERIALS = 156000
    Private Sub frmDelivary_Forcust_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtDyeDate.Text = Today
        txtFromDate.Text = Today
        txtTodate.Text = Today

        Call Load_Gride_SalesOrder()
        Call Load_Material()
        Call Load_Merch()
        Call Load_Retailer()
    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Function Load_Gride_SalesOrder()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer = MakeDataTable_Delivary_Quatation()
        UltraGrid3.DataSource = c_dataCustomer
        With UltraGrid3
            .DisplayLayout.Bands(0).Columns(0).Width = 50
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 70
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(3).Width = 130
            .DisplayLayout.Bands(0).Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(8).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(9).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(12).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(10).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(11).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(13).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(14).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(15).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(16).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '  .DisplayLayout.Bands(0).Columns(17).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            ' .DisplayLayout.Bands(0).Columns(18).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(19).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '  .DisplayLayout.Bands(0).Columns(20).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(3).Width = 60
            '.DisplayLayout.Bands(0).Columns(5).Width = 60
            '.DisplayLayout.Bands(0).Columns(8).Width = 60
            '.DisplayLayout.Bands(0).Columns(7).Width = 70
            '.DisplayLayout.Bands(0).Columns(9).Width = 60

        End With
    End Function

    Function MakeDataTable_Delivary_Quatation() As DataTable
        Dim I As Integer
        Dim X As Integer
        Dim _Lastweek As Integer


        ' MsgBox(DatePart("ww", Today))
        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Sales Order", GetType(String))
        dataTable.Columns.Add(colWork)


        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Line Item", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Material", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Material Description", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Del.qty as today", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        Dim _Del As String

        If Month(Today) = 1 Then
            _Del = "B/T/Del December"
        Else

            _Del = "B/T/Del " & MonthName(Month(Today) - 1)
        End If
        colWork = New DataColumn(_Del, GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        _Del = "B/T/Del " & MonthName(Month(Today))
        colWork = New DataColumn(_Del, GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        'If Month(Today) = 12 Then
        '    _Del = "B/T/Del January"
        'Else
        '    _Del = "B/T/Del " & MonthName(Month(Today) + 1)
        'End If
        'colWork = New DataColumn(_Del, GetType(String))
        'colWork.MaxLength = 250
        'dataTable.Columns.Add(colWork)
        'colWork.ReadOnly = True

        If Month(Today) = 12 Then
            _Del = "1st Wk of ( " & MonthName(1) & ")"
        Else
            _Del = "1st Wk of ( " & MonthName(Month(Today) + 1) & ")"
        End If
        _TobeDis = _Del
        colWork = New DataColumn(_Del, GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        If Month(Today) = 12 Then
            _Del = "2nd Wk of (" & MonthName(1) & ")"
        Else
            _Del = "2nd Wk of (" & MonthName(Month(Today) + 1) & ")"
        End If
        _TobeDis1 = _Del

        colWork = New DataColumn(_Del, GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True


        colWork = New DataColumn("Delivary from Stock", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        _Del = "Pro.to del " & MonthName(Month(Today))
        _ProDis1 = _Del
        colWork = New DataColumn(_Del, GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        '   colWork.ReadOnly = True

        If Month(Today) = 12 Then
            _Del = "Pro.to del January"
        Else
            _Del = "Pro.to del " & MonthName(Month(Today) + 1)
        End If

        _ProDis2 = _Del
        colWork = New DataColumn(_Del, GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        ' colWork.ReadOnly = True

        'If Month(Today) + 2 = 12 Then
        '    _Del = "Pro.to del January"
        'Else
        '    _Del = "Pro.to del " & MonthName(Month(Today) + 2)
        'End If
        '_ProDis3 = _Del
        'colWork = New DataColumn(_Del, GetType(String))
        'colWork.MaxLength = 250
        'dataTable.Columns.Add(colWork)
        ''colWork.ReadOnly = True


        colWork = New DataColumn("Possi Qty From WIP", GetType(String))
        '  colWork.MaxLength = 70
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Excess", GetType(String))
        '  colWork.MaxLength = 70
        dataTable.Columns.Add(colWork)
        '---------------------------------------------------------------------------------

        'If Month(Today) = 1 Then
        '    _Del = "Stock(December)"
        'Else

        '    _Del = "Stock(" & MonthName(Month(Today) - 1) & ")"
        'End If
        '_StockDis1 = _Del
        'colWork = New DataColumn(_Del, GetType(String))
        'colWork.MaxLength = 250
        'dataTable.Columns.Add(colWork)
        'colWork.ReadOnly = True


        '_Del = "Stock(" & MonthName(Month(Today)) & ")"
        'colWork = New DataColumn(_Del, GetType(String))
        'colWork.MaxLength = 250
        'dataTable.Columns.Add(colWork)
        'colWork.ReadOnly = True
        '_StockDis2 = _Del

        ''If Month(Today) = 12 Then
        ''    _Del = "Stock(January)"
        ''Else
        ''    _Del = "Stock(" & MonthName(Month(Today) + 1) & ")"
        ''End If
        ''colWork = New DataColumn(_Del, GetType(String))
        ''colWork.MaxLength = 250
        ''dataTable.Columns.Add(colWork)
        ''colWork.ReadOnly = True

        'If Month(Today) = 12 Then
        '    _Del = "1st Wk (Stock)"
        'Else
        '    _Del = "1st Wk (Stock) " '& MonthName(Month(Today) + 1)
        'End If
        '_StockWK1 = _Del

        'colWork = New DataColumn(_Del, GetType(String))
        'colWork.MaxLength = 250
        'dataTable.Columns.Add(colWork)
        'colWork.ReadOnly = True

        'If Month(Today) = 12 Then
        '    _Del = "2nd Wk (Stock)"
        'Else
        '    _Del = "2nd Wk (Stock) " ' & MonthName(Month(Today) + 1)
        'End If

        '_StockWK2 = _Del

        'colWork = New DataColumn(_Del, GetType(String))
        'colWork.MaxLength = 250
        'dataTable.Columns.Add(colWork)
        'colWork.ReadOnly = True

        'If Month(Today) = 2 Then
        '    _Del = "WIP(December)"
        'ElseIf Month(Today) = 1 Then
        '    _Del = "WIP(Nov)"
        'Else

        '    _Del = "WIP(" & MonthName(Month(Today) - 2) & ")"
        'End If



        'If Month(Today) = 1 Then
        '    _Del = "WIP(December)"
        'Else
        '    _Del = "WIP(" & MonthName(Month(Today) - 1) & ")"
        'End If




        _Del = "WIP(" & MonthName(Month(Today)) & ")"
        _Del = "WIP(Exam)"
        colWork = New DataColumn(_Del, GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        _Del = "WIP(Finishing)"

        colWork = New DataColumn(_Del, GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        _Del = "WIP(Dyeing)"
        colWork = New DataColumn(_Del, GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        'If Month(Today) = 1 Then
        '    _Del = "1st Wk (WIP)"
        'Else
        '    _Del = "1st Wk of " & MonthName(Month(Today) + 1) & "(WIP)"
        '    _Del = "1st Wk (WIP)"
        'End If

        'If Month(Today) = 1 Then
        '    _Del = "2nd Wk(WIP)"
        'Else
        '    _Del = "2nd Wk of " & MonthName(Month(Today) + 1) & "(WIP)"
        '    _Del = "2nd Wk(WIP)"
        'End If

        ''colWork = New DataColumn("#", GetType(String))
        ' ''  colWork.MaxLength = 70
        ''dataTable.Columns.Add(colWork)

        'If Month(Today) = 2 Then
        '    _Del = "To be dyed(December)"
        'ElseIf Month(Today) = 1 Then
        '    _Del = "To be dyed(Nov)"
        'Else

        '    _Del = "To be dyed(" & MonthName(Month(Today) - 2) & ")"
        'End If

        '_Del = "To be dyed"
        'colWork = New DataColumn(_Del, GetType(String))
        'colWork.MaxLength = 250
        'dataTable.Columns.Add(colWork)
        'colWork.ReadOnly = True

        'If Month(Today) = 1 Then
        '    _Del = "To be dyed(December)"
        'Else
        '    _Del = "To be dyed(" & MonthName(Month(Today) - 1) & ")"
        'End If
        'colWork = New DataColumn(_Del, GetType(String))
        'colWork.MaxLength = 250
        'dataTable.Columns.Add(colWork)
        'colWork.ReadOnly = True


        '_Del = "To be dyed(" & MonthName(Month(Today)) & ")"
        'colWork = New DataColumn(_Del, GetType(String))
        'colWork.MaxLength = 250
        'dataTable.Columns.Add(colWork)
        'colWork.ReadOnly = True

        'If Month(Today) = 1 Then
        '    _Del = "1st Wk of January(T/D)"
        'Else
        '    _Del = "1st Wk of " & MonthName(Month(Today) + 1) & "(T/D)"
        'End If

        'If Month(Today) = 1 Then
        '    _Del = "2nd Wk of January(T/D)"
        'Else
        '    _Del = "2nd Wk of " & MonthName(Month(Today) + 1) & "(T/D)"
        'End If
        Return dataTable
    End Function

    Function Load_Material()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Dim I As Integer

        Try

            Sql = "SELECT M07Material as [Material] FROM M07TobeDelivered  UNION SELECT M06Material FROM M06Delivary_Qty"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboMaterial
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 175
            End With

            DBEngin.CloseConnection(con)
            con.ConnectionString = ""


        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
            End If
        End Try

    End Function

    Function Load_Merch()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Dim I As Integer

        Try

            Sql = "SELECT Merchant as [Merchant] FROM OTD_SMS group by Merchant"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboMerch
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 175
            End With

            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            con.close()


        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
            End If
        End Try

    End Function

    Function Load_Retailer()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Dim I As Integer

        Try

            Sql = "SELECT  M07Retailer as [Retailer] FROM View_DelivaryForcus group by  M07Retailer"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboRetailer
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 175
            End With

            DBEngin.CloseConnection(con)
            con.ConnectionString = ""

            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
            End If
        End Try

    End Function

    Function Find_Data()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Dim I As Integer
        Dim _Total_PossibleQTY As Double
        Dim _Total_PossibleQTY1 As Double
        Dim _TotalStock As Double
        Dim _DelQty As Double
        Dim _BTQty01 As Double
        Dim _TotalDel_Qty As Double

        Dim _BT1 As String
        Dim _BTQty02 As Double
        Dim _BT2 As String

        Dim _BTQty03 As Double
        Dim _BT3 As String
        Dim _Stok1 As Double
        Dim _Stock2 As Double

        Dim _StokW1 As Double
        Dim _StockW2 As Double
        Dim M08 As DataSet
        Dim _RowIndex As Integer
        Dim _StFirstMterial As String
        Dim _Fromdate As Date
        Dim _Todate As Date

        Try

            _Fromdate = txtFromDate.Text

            _Fromdate = _Fromdate
            Dim thisCulture1 = Globalization.CultureInfo.CurrentCulture
            Dim dayOfWeek1 As DayOfWeek = thisCulture1.Calendar.GetDayOfWeek(_Fromdate)
            Dim dayName1 As String = thisCulture1.DateTimeFormat.GetDayName(dayOfWeek1)

            Dim daysInFeb As Integer = System.DateTime.DaysInMonth(Year(_Fromdate), Month(_Fromdate))
            _Todate = CDate(_Fromdate).AddDays(+daysInFeb)
            _Todate = _Todate.AddDays(-1)


            Call Load_Gride_SalesOrder()
            If Trim(cboMaterial.Text) <> "" Then
                Sql = "select M07Sales_Order,M07Line_Item,max(M07Material) as M07Material,max(M07Met_Dis) as M07Met_Dis ,sum(M07Qty_Mtr) as M07Qty_Mtr from View_DelivaryForcus where M07Material='" & Trim(cboMaterial.Text) & "' and M07Date <= '" & txtTodate.Text & "' group by M07Sales_Order,M07Line_Item order by M07Material,M07Sales_Order,M07Line_Item"

            ElseIf cboMerch.Text <> "" Then
                Sql = "select M07Sales_Order,M07Line_Item,max(M07Material) as M07Material,max(M07Met_Dis) as M07Met_Dis ,sum(M07Qty_Mtr) as M07Qty_Mtr from View_DelivaryForcus  inner join OTD_SMS on Sales_Order=M07Sales_Order and M07Line_Item=Line_Item where  M07Date <= '" & txtTodate.Text & "' and M07Merchant='" & cboMerch.Text & "' group by M07Sales_Order,M07Line_Item order by M07Material,M07Sales_Order,M07Line_Item"

            ElseIf cboRetailer.Text <> "" Then
                Sql = "select M07Sales_Order,M07Line_Item,max(M07Material) as M07Material,max(M07Met_Dis) as M07Met_Dis ,sum(M07Qty_Mtr) as M07Qty_Mtr from View_DelivaryForcus  where  M07Retailer='" & cboRetailer.Text & "' group by M07Sales_Order,M07Line_Item order by M07Material,M07Sales_Order,M07Line_Item"

            Else
                Sql = "select M07Sales_Order,M07Line_Item,max(M07Material) as M07Material,max(M07Met_Dis) as M07Met_Dis ,sum(M07Qty_Mtr) as M07Qty_Mtr from View_DelivaryForcus where M07Date <= '" & txtTodate.Text & "' group by M07Sales_Order,M07Line_Item order by M07Material,M07Sales_Order,M07Line_Item"
                ' Sql = "select M07Sales_Order,M07Line_Item,max(M07Material) as M07Material,max(M07Met_Dis) as M07Met_Dis ,sum(M07Qty_Mtr) as M07Qty_Mtr from View_DelivaryForcus group by M07Sales_Order,order by M07Material"
            End If
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            I = 0
            _DelQty = 0
            _BTQty01 = 0
            _BTQty02 = 0
            _BTQty03 = 0
            _Stock2 = 0
            _Stok1 = 0
            _StokW1 = 0
            _StockW2 = 0
            _RowIndex = 0
            _TobeWkQty = 0
            _TobeWKQty1 = 0
            _WIPDye = 0
            _WIPExam = 0
            _WIPFinishing = 0
            _Promiss_To_delTot = 0
            _Total_PossibleQTY1 = 0
            _TotalDel_Qty = 0
            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer.NewRow
                Dim Value As Double
                Dim _Dis As String
                Dim _From As Date
                Dim _To As Date
                Dim EndDate As DateTime

                _Total_PossibleQTY = 0

                'If M01.Tables(0).Rows(I)("M07Sales_Order") = "4004760" Then
                '    ' MsgBox("")
                'End If
                If I = 8 Then
                    '   MsgBox("")
                End If
                'Sql = "select M06Sales_Order,M06Line_Item,sum(M06D_Qty_Mtr) as M06D_Qty_Mtr from M06Delivary_Qty where M06Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and  M06Sales_Order='" & Trim(M01.Tables(0).Rows(I)("M07Sales_Order")) & "' and M06Line_Item='" & Trim(M01.Tables(0).Rows(I)("M07Line_Item")) & "'  group by M06Sales_Order,M06Line_Item"
                'dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
                'If isValidDataset(dsUser) Then
                newRow("Sales Order") = M01.Tables(0).Rows(I)("M07Sales_Order")
                newRow("Line Item") = M01.Tables(0).Rows(I)("M07Line_Item")
                newRow("Material") = M01.Tables(0).Rows(I)("M07Material")
                newRow("Material Description") = M01.Tables(0).Rows(I)("M07Met_Dis")

               
                Sql = "select M06Sales_Order,M06Line_Item,sum(M06D_Qty_Mtr) as M06D_Qty_Mtr from M06Delivary_Qty where M06Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and  M06Sales_Order='" & Trim(M01.Tables(0).Rows(I)("M07Sales_Order")) & "' and M06Line_Item='" & Trim(M01.Tables(0).Rows(I)("M07Line_Item")) & "'  group by M06Sales_Order,M06Line_Item"
                dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(dsUser) Then
                    Value = dsUser.Tables(0).Rows(0)("M06D_Qty_Mtr")
                    _DelQty = _DelQty + dsUser.Tables(0).Rows(0)("M06D_Qty_Mtr")
                    _Dis = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    _Dis = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0}", Value))
                    newRow("Del.qty as today") = _Dis
                End If

                'STOCK

                'Sql = "SELECT SUM(M08Qty_Mtr) AS M08Qty_Mtr FROM M08Stock WHERE M08Sales_Order='" & Trim(M01.Tables(0).Rows(I)("M07Sales_Order")) & "' AND M08Line_Item='" & Trim(M01.Tables(0).Rows(I)("M07Line_Item")) & "' AND M08Meterial='" & M01.Tables(0).Rows(I)("M07Material") & "' and M08Location in ('2060','2059') GROUP BY M08Sales_Order,M08Line_Item"
                'dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
                'If isValidDataset(dsUser) Then
                '    _TotalStock = dsUser.Tables(0).Rows(0)("M08Qty_Mtr")
                'End If


                Dim L_DateofMonth As Integer

                _From = CDate(Today)
                _From = Month(_From) & "/1/" & Year(_From)
                If Month(_From) = 1 Then
                    _From = "12/1/" & Year(_From) - 1
                Else
                    _From = Month(_From) - 1 & "/1/" & Year(_From)
                End If
                EndDate = _From.AddDays(DateTime.DaysInMonth(_From.Year, _From.Month) - 1)
                L_DateofMonth = Microsoft.VisualBasic.Day(EndDate)
                L_DateofMonth = L_DateofMonth - 1
                _To = _From.AddDays(+L_DateofMonth)
                Dim _Del As String
                Sql = "select sum(M07Qty_Mtr) as M07Qty_Mtr from M07TobeDelivered  where  M07Date <= '" & _To & "' and  M07Sales_Order='" & Trim(M01.Tables(0).Rows(I)("M07Sales_Order")) & "' and M07Line_Item='" & Trim(M01.Tables(0).Rows(I)("M07Line_Item")) & "'   group by M07Sales_Order,M07Line_Item"
                dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(dsUser) Then
                    If Month(Today) = 1 Then
                        _Del = "B/T/Del Des"

                    Else
                        _Del = "B/T/Del " & MonthName(Month(Today) - 1)
                    End If
                    _BT1 = _Del

                    _BTQty01 = _BTQty01 + dsUser.Tables(0).Rows(0)("M07Qty_Mtr")

                    Value = dsUser.Tables(0).Rows(0)("M07Qty_Mtr")
                    _Dis = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    _Dis = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0}", Value))
                    newRow(_Del) = _Dis

                    '    If Month(Today) = 1 Then
                    '        _Del = "Stock(Des)"

                    '    Else
                    '        _Del = "Stock(" & MonthName(Month(Today) - 1) & ")"
                    '    End If

                    '    If dsUser.Tables(0).Rows(0)("M07Qty_Mtr") < _TotalStock Then
                    '        If dsUser.Tables(0).Rows(0)("M07Qty_Mtr") = 0 Then
                    '        Else
                    '            _TotalStock = _TotalStock - dsUser.Tables(0).Rows(0)("M07Qty_Mtr")
                    '            Value = dsUser.Tables(0).Rows(0)("M07Qty_Mtr")
                    '            _Dis = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    '            _Dis = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0}", Value))
                    '            newRow(_Del) = _Dis
                    '            _Stok1 = _Stok1 + Value

                    '            _Total_PossibleQTY = _Total_PossibleQTY + Value
                    '        End If
                    '    ElseIf dsUser.Tables(0).Rows(0)("M07Qty_Mtr") >= _TotalStock Then
                    '        _Total_PossibleQTY = _Total_PossibleQTY + _TotalStock
                    '        _Stok1 = _Stok1 + _TotalStock
                    '        Value = _TotalStock
                    '        _Dis = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    '        _Dis = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0}", Value))
                    '        newRow(_Del) = _Dis
                    '        _TotalStock = 0
                    '    End If
                End If

                _From = Month(Today) & "/1/" & Year(Today)
                EndDate = _From.AddDays(DateTime.DaysInMonth(_From.Year, _From.Month) - 1)
                'L_DateofMonth = Microsoft.VisualBasic.Day(EndDate)
                'L_DateofMonth = L_DateofMonth - 1
                '_To = _From.AddDays(+L_DateofMonth)

                Sql = "select sum(M07Qty_Mtr) as M07Qty_Mtr from M07TobeDelivered  where M07Sales_Order='" & Trim(M01.Tables(0).Rows(I)("M07Sales_Order")) & "' and M07Line_Item='" & Trim(M01.Tables(0).Rows(I)("M07Line_Item")) & "' and M07Date between '" & _From & "' and '" & EndDate & "' group by M07Sales_Order,M07Line_Item"
                dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(dsUser) Then

                    _Del = "B/T/Del " & MonthName(Month(Today))

                    Value = dsUser.Tables(0).Rows(0)("M07Qty_Mtr")
                    _Dis = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    _Dis = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0}", Value))
                    newRow(_Del) = _Dis

                    _BT2 = _Del

                    _BTQty02 = _BTQty02 + dsUser.Tables(0).Rows(0)("M07Qty_Mtr")

                    '    _Del = "Stock(" & MonthName(Month(Today)) & ")"
                    '    If dsUser.Tables(0).Rows(0)("M07Qty_Mtr") <= _TotalStock Then
                    '        If dsUser.Tables(0).Rows(0)("M07Qty_Mtr") = 0 Then
                    '        Else
                    '            _TotalStock = _TotalStock - dsUser.Tables(0).Rows(0)("M07Qty_Mtr")
                    '            Value = dsUser.Tables(0).Rows(0)("M07Qty_Mtr")
                    '            _Dis = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    '            _Dis = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0}", Value))
                    '            newRow(_Del) = _Dis
                    '            _Stock2 = _Stock2 + dsUser.Tables(0).Rows(0)("M07Qty_Mtr")
                    '            _Total_PossibleQTY = _Total_PossibleQTY + dsUser.Tables(0).Rows(0)("M07Qty_Mtr")
                    '        End If
                    '    ElseIf dsUser.Tables(0).Rows(0)("M07Qty_Mtr") > _TotalStock Then
                    '        Value = _TotalStock
                    '        _Total_PossibleQTY = _Total_PossibleQTY + Value
                    '        _Stock2 = _Stock2 + _TotalStock
                    '        _Dis = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    '        _Dis = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0}", Value))
                    '        newRow(_Del) = _Dis
                    '        _TotalStock = 0
                    '    End If
                End If

                _From = Today
                If Month(_From) = 12 Then
                    _From = "1/1/" & Year(_From) + 1
                Else
                    _From = Month(Today) + 1 & "/1/" & Year(Today)
                End If

                EndDate = _From.AddDays(DateTime.DaysInMonth(_From.Year, _From.Month) - 1)
                'L_DateofMonth = Microsoft.VisualBasic.Day(EndDate)
                'L_DateofMonth = L_DateofMonth - 1
                '_To = _From.AddDays(+L_DateofMonth)

                Sql = "select sum(M07Qty_Mtr) as M07Qty_Mtr from M07TobeDelivered  where M07Sales_Order='" & Trim(M01.Tables(0).Rows(I)("M07Sales_Order")) & "' and M07Line_Item='" & Trim(M01.Tables(0).Rows(I)("M07Line_Item")) & "' and M07Date >= '" & _From & "' group by M07Sales_Order,M07Line_Item"
                dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(dsUser) Then

                    '  _Del = "B/T/Del " & MonthName(Month(_From))

                    Value = dsUser.Tables(0).Rows(0)("M07Qty_Mtr")
                    '_Dis = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    '_Dis = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0}", Value))
                    'newRow(_Del) = _Dis
                    '' End If

                    '_BT3 = _Del

                    '_BTQty03 = _BTQty03 + dsUser.Tables(0).Rows(0)("M07Qty_Mtr")

                    If Value > 0 Then
                        Dim _wkFIRST As Date
                        Dim _wkTo As Date
                        _wkFIRST = Month(_From) & "/1/" & Year(_From)

                        Dim thisCulture = Globalization.CultureInfo.CurrentCulture
                        Dim dayOfWeek As DayOfWeek = thisCulture.Calendar.GetDayOfWeek(_wkFIRST)
                        Dim dayName As String = thisCulture.DateTimeFormat.GetDayName(dayOfWeek)

                        If dayName = "Sunday" Then
                            _wkTo = CDate(_wkFIRST).AddDays(+6)
                        ElseIf dayName = "Tuesday" Then
                            _wkTo = CDate(_wkFIRST).AddDays(+5)
                        ElseIf dayName = "Wednesday" Then
                            _wkTo = CDate(_wkFIRST).AddDays(+4)
                        ElseIf dayName = "Thuesday" Then
                            _wkTo = CDate(_wkFIRST).AddDays(+3)
                        ElseIf dayName = "Friday" Then
                            _wkTo = CDate(_wkFIRST).AddDays(+2)
                        ElseIf dayName = "Saturday" Then
                            _wkTo = CDate(_wkFIRST).AddDays(+1)
                        ElseIf dayName = "Monday" Then
                            _wkTo = CDate(_wkFIRST).AddDays(+7)
                        End If
                        '1ST WEEK


                        Sql = "select sum(M07Qty_Mtr) as M07Qty_Mtr from M07TobeDelivered  where M07Sales_Order='" & Trim(M01.Tables(0).Rows(I)("M07Sales_Order")) & "' and M07Line_Item='" & Trim(M01.Tables(0).Rows(I)("M07Line_Item")) & "' and M07Date between '" & _wkFIRST & "' and '" & _wkTo & "' group by M07Sales_Order,M07Line_Item"
                        M08 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                        If isValidDataset(M08) Then
                            _TobeWkQty = _TobeWkQty + M08.Tables(0).Rows(0)("M07Qty_Mtr")

                            Value = M08.Tables(0).Rows(0)("M07Qty_Mtr")
                            _Dis = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                            _Dis = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0}", Value))
                            newRow(_TobeDis) = _Dis
                            _Total_PossibleQTY = _Total_PossibleQTY + Value
                            If dsUser.Tables(0).Rows(0)("M07Qty_Mtr") <= _TotalStock Then
                                If dsUser.Tables(0).Rows(0)("M07Qty_Mtr") = 0 Then
                                Else
                                    _StokW1 = _TotalStock - dsUser.Tables(0).Rows(0)("M07Qty_Mtr")
                                    Value = dsUser.Tables(0).Rows(0)("M07Qty_Mtr")
                                    _Dis = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                                    _Dis = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0}", Value))
                                    newRow(_StockWK1) = _Dis

                                    _Total_PossibleQTY = _Total_PossibleQTY + Value
                                End If
                            ElseIf dsUser.Tables(0).Rows(0)("M07Qty_Mtr") > _TotalStock Then
                                'Value = _TotalStock
                                '_Total_PossibleQTY = _Total_PossibleQTY + Value
                                '_Dis = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                                '_Dis = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0}", Value))
                                'newRow(_StockWK1) = _Dis

                                '_TotalStock = 0
                            End If
                        End If
                        'Sql = "SELECT SUM(M08Qty_Mtr) AS M08Qty_Mtr FROM M08Stock WHERE M08Sales_Order='" & Trim(M01.Tables(0).Rows(I)("M07Sales_Order")) & "' AND M08Line_Item='" & Trim(M01.Tables(0).Rows(I)("M07Line_Item")) & "' AND M08Meterial='" & M01.Tables(0).Rows(I)("M07Material") & "' and M08Location in ('2060','2059') GROUP BY M08Sales_Order,M08Line_Item"
                        'dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
                        'If isValidDataset(dsUser) Then
                        '    _StokW1 = dsUser.Tables(0).Rows(0)("M08Qty_Mtr")
                        'End If

                        _wkFIRST = _wkTo.AddDays(+1)
                        _wkTo = _wkFIRST.AddDays(+6)

                        Sql = "select sum(M07Qty_Mtr) as M07Qty_Mtr from M07TobeDelivered  where M07Sales_Order='" & Trim(M01.Tables(0).Rows(I)("M07Sales_Order")) & "' and M07Line_Item='" & Trim(M01.Tables(0).Rows(I)("M07Line_Item")) & "' and M07Date between '" & _wkFIRST & "' and '" & _wkTo & "' group by M07Sales_Order,M07Line_Item"
                        M08 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                        If isValidDataset(M08) Then
                            _TobeWKQty1 = _TobeWKQty1 + M08.Tables(0).Rows(0)("M07Qty_Mtr")
                            Value = M08.Tables(0).Rows(0)("M07Qty_Mtr")

                            _Dis = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                            _Dis = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0}", Value))
                            newRow(_TobeDis1) = _Dis

                            '    If dsUser.Tables(0).Rows(0)("M07Qty_Mtr") <= _StokW1 Then
                            '        If dsUser.Tables(0).Rows(0)("M07Qty_Mtr") = 0 Then
                            '        Else
                            '            _StockW2 = _StokW1 - dsUser.Tables(0).Rows(0)("M07Qty_Mtr")
                            '            Value = dsUser.Tables(0).Rows(0)("M07Qty_Mtr")
                            '            _Total_PossibleQTY = _Total_PossibleQTY + Value
                            '            _Dis = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                            '            _Dis = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0}", Value))
                            '            newRow(_StockWK2) = _Dis
                            '        End If
                            '    ElseIf dsUser.Tables(0).Rows(0)("M07Qty_Mtr") > _StokW1 Then
                            '        Value = _StokW1
                            '        _Total_PossibleQTY = _Total_PossibleQTY + Value
                            '        _Dis = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                            '        _Dis = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0}", Value))
                            '        newRow(_StockWK1) = _Dis
                            '        '_TotalStock = 0
                            '    End If
                            'End If

                        End If
                        '==================================================================================
                        '_Del = "Stock(" & MonthName(Month(_From)) & ")"
                        'If dsUser.Tables(0).Rows(0)("M07Qty_Mtr") <= _TotalStock Then
                        '    If dsUser.Tables(0).Rows(0)("M07Qty_Mtr") = 0 Then
                        '    Else
                        '        _TotalStock = _TotalStock - dsUser.Tables(0).Rows(0)("M07Qty_Mtr")
                        '        Value = dsUser.Tables(0).Rows(0)("M07Qty_Mtr")
                        '        _Dis = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        '        _Dis = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0}", Value))
                        '        newRow(_Del) = _Dis
                        '    End If
                        'ElseIf dsUser.Tables(0).Rows(0)("M07Qty_Mtr") > _TotalStock Then
                        '    'Value = _TotalStock
                        '    '_Dis = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        '    '_Dis = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0}", Value))
                        '    'newRow(_Del) = _Dis
                        '    '_TotalStock = 0
                        'End If

                        '===================================================================
                        '1ST WEEK
                        'Sql = "SELECT SUM(M08Qty_Mtr) AS M08Qty_Mtr FROM M08Stock WHERE M08Sales_Order='" & Trim(M01.Tables(0).Rows(I)("M07Sales_Order")) & "' AND M08Line_Item='" & Trim(M01.Tables(0).Rows(I)("M07Line_Item")) & "' AND M08Meterial='" & M01.Tables(0).Rows(I)("M07Material") & "' and M08Location in ('2060','2059') GROUP BY M08Sales_Order,M08Line_Item"
                        'dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
                        'If isValidDataset(dsUser) Then
                        '    _StokW1 = dsUser.Tables(0).Rows(0)("M08Qty_Mtr")
                        'End If

                    End If
                End If
                '===================================================================================
                Dim _Exam As Double
                Dim _Finishing As Double
                Dim _Dye As Double

                _Exam = 0
                _Dye = 0
                _Finishing = 0

                'WIP
                If chkCus.Checked = True Then
                    Sql = "select sum(M09Qty_Mtr) as M09Qty_Mtr from M09ZPL_ORDER inner join FR_Update on M09BatchNo=Batch_No where M09Oredr_Type='Dyeing' and M09Meterial='" & Trim(M01.Tables(0).Rows(I)("M07Material")) & "' and Dye_Pln_Date<= '" & txtDyeDate.Text & "' group by  M09Oredr_Type"
                    M08 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(M08) Then
                        If Trim(_StFirstMterial) = Trim(M01.Tables(0).Rows(I)("M07Material")) Then

                        Else
                            Value = M08.Tables(0).Rows(0)("M09Qty_Mtr")
                            _Dye = M08.Tables(0).Rows(0)("M09Qty_Mtr")
                            _Total_PossibleQTY = _Total_PossibleQTY + Value
                            _Dis = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                            _Dis = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0}", Value))
                            newRow("WIP(Dyeing)") = _Dis

                            _WIPDye = _WIPDye + M08.Tables(0).Rows(0)("M09Qty_Mtr")
                            '_StFirstMterial = Trim(M01.Tables(0).Rows(0)("M07Material"))
                            'UltraGrid3.Rows(_RowIndex).Cells(1).Appearance.BackColor = Color.DarkBlue
                            'UltraGrid3.Rows(_RowIndex).Cells(2).Appearance.BackColor = Color.DarkBlue
                            'UltraGrid3.Rows(_RowIndex).Cells(3).Appearance.BackColor = Color.DarkBlue
                            'UltraGrid3.Rows(_RowIndex).Cells(4).Appearance.BackColor = Color.DarkBlue
                            'UltraGrid3.Rows(_RowIndex).Cells(5).Appearance.BackColor = Color.DarkBlue
                            'UltraGrid3.Rows(_RowIndex).Cells(6).Appearance.BackColor = Color.DarkBlue
                            'UltraGrid3.Rows(_RowIndex).Cells(7).Appearance.BackColor = Color.DarkBlue
                            'UltraGrid3.Rows(_RowIndex).Cells(8).Appearance.BackColor = Color.DarkBlue
                            'UltraGrid3.Rows(_RowIndex).Cells(9).Appearance.BackColor = Color.DarkBlue
                            'UltraGrid3.Rows(_RowIndex).Cells(10).Appearance.BackColor = Color.DarkBlue
                            'UltraGrid3.Rows(_RowIndex).Cells(11).Appearance.BackColor = Color.DarkBlue
                            'UltraGrid3.Rows(_RowIndex).Cells(12).Appearance.BackColor = Color.DarkBlue
                            'UltraGrid3.Rows(_RowIndex).Cells(13).Appearance.BackColor = Color.DarkBlue


                        End If
                    End If
                End If
                '---------------------------------------------------------------------------------------------
                Sql = "select sum(M09Qty_Mtr) as M09Qty_Mtr from M09ZPL_ORDER  where M09Oredr_Type='Exam' and M09Meterial='" & Trim(M01.Tables(0).Rows(I)("M07Material")) & "' group by  M09Oredr_Type"
                M08 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(M08) Then
                    If Trim(_StFirstMterial) = Trim(M01.Tables(0).Rows(I)("M07Material")) Then

                    Else
                        Value = M08.Tables(0).Rows(0)("M09Qty_Mtr")
                        _Exam = M08.Tables(0).Rows(0)("M09Qty_Mtr")
                        _Total_PossibleQTY = _Total_PossibleQTY + Value
                        _Dis = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        _Dis = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0}", Value))
                        newRow("WIP(Exam)") = _Dis

                        _WIPExam = _WIPExam + M08.Tables(0).Rows(0)("M09Qty_Mtr")
                        ' _StFirstMterial = Trim(M01.Tables(0).Rows(0)("M07Material"))
                        'UltraGrid3.Rows(_RowIndex).Cells(1).Appearance.BackColor = Color.DarkBlue
                        'UltraGrid3.Rows(_RowIndex).Cells(2).Appearance.BackColor = Color.DarkBlue
                        'UltraGrid3.Rows(_RowIndex).Cells(3).Appearance.BackColor = Color.DarkBlue
                        'UltraGrid3.Rows(_RowIndex).Cells(4).Appearance.BackColor = Color.DarkBlue
                        'UltraGrid3.Rows(_RowIndex).Cells(5).Appearance.BackColor = Color.DarkBlue
                        'UltraGrid3.Rows(_RowIndex).Cells(6).Appearance.BackColor = Color.DarkBlue
                        'UltraGrid3.Rows(_RowIndex).Cells(7).Appearance.BackColor = Color.DarkBlue
                        'UltraGrid3.Rows(_RowIndex).Cells(8).Appearance.BackColor = Color.DarkBlue
                        'UltraGrid3.Rows(_RowIndex).Cells(9).Appearance.BackColor = Color.DarkBlue
                        'UltraGrid3.Rows(_RowIndex).Cells(10).Appearance.BackColor = Color.DarkBlue
                        'UltraGrid3.Rows(_RowIndex).Cells(11).Appearance.BackColor = Color.DarkBlue
                        'UltraGrid3.Rows(_RowIndex).Cells(12).Appearance.BackColor = Color.DarkBlue
                        'UltraGrid3.Rows(_RowIndex).Cells(13).Appearance.BackColor = Color.DarkBlue


                    End If

                End If

                Sql = "select sum(M09Qty_Mtr) as M09Qty_Mtr from M09ZPL_ORDER  where M09Oredr_Type='Finishing' and M09Meterial='" & Trim(M01.Tables(0).Rows(I)("M07Material")) & "' group by  M09Oredr_Type"
                M08 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(M08) Then
                    If Trim(_StFirstMterial) = Trim(M01.Tables(0).Rows(I)("M07Material")) Then

                    Else
                        Value = M08.Tables(0).Rows(0)("M09Qty_Mtr")
                        _Finishing = M08.Tables(0).Rows(0)("M09Qty_Mtr")
                        _Total_PossibleQTY = _Total_PossibleQTY + Value
                        _Dis = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        _Dis = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0}", Value))
                        newRow("WIP(Finishing)") = _Dis

                        _WIPFinishing = _WIPFinishing + M08.Tables(0).Rows(0)("M09Qty_Mtr")



                    End If

                End If
                _Promiss_To_delTot = _Promiss_To_delTot + _Total_PossibleQTY
                'If _Total_PossibleQTY > 0 Then
                '    Value = _Total_PossibleQTY
                '    ' _Total_PossibleQTY = _Total_PossibleQTY + Value
                '    _Dis = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                '    _Dis = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0}", Value))
                '    newRow("Tot.Possi Qty") = _Dis
                'End If

                If _StFirstMterial = Trim(M01.Tables(0).Rows(I)("M07Material")) Then
                Else
                    Sql = "select SUM(M08Qty_Mtr) as M08Qty_Mtr from M08Stock where M08Meterial='" & Trim(M01.Tables(0).Rows(I)("M07Material")) & "' group by M08Meterial"
                    M08 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(M08) Then
                        ' newRow("Possi Qty From WIP") = M08.Tables(0).Rows(0)("M08Qty_Mtr") + _Exam + _Dye + _Finishing
                        newRow("Possi Qty From WIP") = (_Exam + _Dye + _Finishing)
                        '_Total_PossibleQTY1 = _Total_PossibleQTY1 + (M08.Tables(0).Rows(0)("M08Qty_Mtr") + _Exam + _Dye + _Finishing)
                        _Total_PossibleQTY1 = (_Exam + _Dye + _Finishing)
                    End If

                End If
                Sql = "select * from T08DelayComment where T08Sales_Order='" & Trim(M01.Tables(0).Rows(I)("M07Sales_Order")) & "' and T08Line_Item='" & Trim(M01.Tables(0).Rows(I)("M07Line_Item")) & "' and T08Material='" & Trim(M01.Tables(0).Rows(I)("M07Material")) & "' and T08Date between '" & _Fromdate & "' and '" & _Todate & "' order by T08Date  DESC"
                M08 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(M08) Then
                    If M08.Tables(0).Rows(0)("T08Pro_Qty1") > 0 Then
                        Value = M08.Tables(0).Rows(0)("T08Pro_Qty1")
                        ' _Total_PossibleQTY = _Total_PossibleQTY + Value
                        _Dis = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        _Dis = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0}", Value))
                        newRow(_ProDis1) = _Dis
                    End If

                    If M08.Tables(0).Rows(0)("T08Pro_Qty2") > 0 Then
                        Value = M08.Tables(0).Rows(0)("T08Pro_Qty2")
                        ' _Total_PossibleQTY = _Total_PossibleQTY + Value
                        _Dis = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        _Dis = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0}", Value))
                        newRow(_ProDis2) = _Dis
                    End If

                End If

                Dim aDate As DateTime
                Dim _Formdate1 As Date
                Dim _Lastdate As Date

                aDate = Month(txtFromDate.Text) & "/1/" & Year(txtFromDate.Text)
                ' _Formdate1 = aDate
                EndDate = aDate.AddDays(DateTime.DaysInMonth(aDate.Year, aDate.Month) - 1)
                '  MsgBox(MonthName(T01.Tables(0).Rows(i)("T01month"), True))
                L_DateofMonth = Microsoft.VisualBasic.Day(EndDate)
                _Lastdate = Month(txtFromDate.Text) & "/" & L_DateofMonth & "/" & Year(txtFromDate.Text)

                '  If cboMerch.Text <> "" Then
                '     Sql = "select sum(Qty) as Qty from View_FG_StockComment where M08Sales_Order='" & M01.Tables(0).Rows(I)("M07Sales_Order") & "' and M08Line_Item='" & M01.Tables(0).Rows(I)("M07Line_Item") & "' and d_Date between '" & aDate & "' and '" & _Lastdate & "' and M08Merchant='" & cboMaterial.Text group by M08Sales_Order,M08Line_Item"
                'Else
                Sql = "select sum(Qty) as Qty from View_FG_StockComment where M08Sales_Order='" & M01.Tables(0).Rows(I)("M07Sales_Order") & "' and M08Line_Item='" & M01.Tables(0).Rows(I)("M07Line_Item") & "' and d_Date between '" & aDate & "' and '" & _Lastdate & "' group by M08Sales_Order,M08Line_Item"
                'End If
                dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(dsUser) Then
                    Value = dsUser.Tables(0).Rows(0)("Qty")
                    _TotalDel_Qty = _TotalDel_Qty + Value
                    _Dis = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    _Dis = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0}", Value))
                    newRow("Delivary from Stock") = _Dis
                End If




                c_dataCustomer.Rows.Add(newRow)

                UltraGrid3.Rows(_RowIndex).Cells(4).Appearance.BackColor = Color.DarkOrange
                UltraGrid3.Rows(_RowIndex).Cells(10).Appearance.BackColor = Color.Green
                UltraGrid3.Rows(_RowIndex).Cells(11).Appearance.BackColor = Color.Green
                UltraGrid3.Rows(_RowIndex).Cells(12).Appearance.BackColor = Color.Yellow

                ' UltraGrid3.Rows(_RowIndex).Cells(18).Appearance.

                If Trim(_StFirstMterial) = Trim(M01.Tables(0).Rows(I)("M07Material")) Then
                Else
                    _StFirstMterial = Trim(M01.Tables(0).Rows(I)("M07Material"))

                    'If Val(UltraGrid3.Rows(_RowIndex).Cells(10).Value) > 0 Or Val(UltraGrid3.Rows(_RowIndex).Cells(19).Value) > 0 Or Val(UltraGrid3.Rows(_RowIndex).Cells(18).Value) > 0 Then
                    ' If Trim(UltraGrid3.Rows(_RowIndex).Cells(18).Text) <> "" Then 'Or Trim(UltraGrid3.Rows(_RowIndex).Cells(19).Text) <> "" Or Trim(UltraGrid3.Rows(_RowIndex).Cells(20).Text) Then
                    UltraGrid3.Rows(_RowIndex).Cells(0).Appearance.BackColor = Color.LightGreen
                    UltraGrid3.Rows(_RowIndex).Cells(1).Appearance.BackColor = Color.LightGreen
                    UltraGrid3.Rows(_RowIndex).Cells(2).Appearance.BackColor = Color.LightGreen
                    UltraGrid3.Rows(_RowIndex).Cells(3).Appearance.BackColor = Color.LightGreen
                    UltraGrid3.Rows(_RowIndex).Cells(4).Appearance.BackColor = Color.LightGreen
                    UltraGrid3.Rows(_RowIndex).Cells(5).Appearance.BackColor = Color.LightGreen
                    UltraGrid3.Rows(_RowIndex).Cells(6).Appearance.BackColor = Color.LightGreen
                    UltraGrid3.Rows(_RowIndex).Cells(7).Appearance.BackColor = Color.LightGreen
                    UltraGrid3.Rows(_RowIndex).Cells(8).Appearance.BackColor = Color.LightGreen
                    UltraGrid3.Rows(_RowIndex).Cells(9).Appearance.BackColor = Color.LightGreen
                    UltraGrid3.Rows(_RowIndex).Cells(10).Appearance.BackColor = Color.LightGreen
                    UltraGrid3.Rows(_RowIndex).Cells(11).Appearance.BackColor = Color.LightGreen
                    UltraGrid3.Rows(_RowIndex).Cells(12).Appearance.BackColor = Color.LightGreen
                    UltraGrid3.Rows(_RowIndex).Cells(13).Appearance.BackColor = Color.LightGreen
                    UltraGrid3.Rows(_RowIndex).Cells(14).Appearance.BackColor = Color.LightGreen
                    UltraGrid3.Rows(_RowIndex).Cells(15).Appearance.BackColor = Color.LightGreen
                    UltraGrid3.Rows(_RowIndex).Cells(16).Appearance.BackColor = Color.LightGreen
                    '  UltraGrid3.Rows(_RowIndex).Cells(17).Appearance.BackColor = Color.LightGreen
                    ' UltraGrid3.Rows(_RowIndex).Cells(18).Appearance.BackColor = Color.LightGreen
                    'UltraGrid3.Rows(_RowIndex).Cells(19).Appearance.BackColor = Color.LightGreen
                    'UltraGrid3.Rows(_RowIndex).Cells(20).Appearance.BackColor = Color.LightGreen
                    '    ElseIf Trim(UltraGrid3.Rows(_RowIndex).Cells(19).Text) <> "" Then

                    '    UltraGrid3.Rows(_RowIndex).Cells(0).Appearance.BackColor = Color.LightGreen
                    '    UltraGrid3.Rows(_RowIndex).Cells(1).Appearance.BackColor = Color.LightGreen
                    '    UltraGrid3.Rows(_RowIndex).Cells(2).Appearance.BackColor = Color.LightGreen
                    '    UltraGrid3.Rows(_RowIndex).Cells(3).Appearance.BackColor = Color.LightGreen
                    '    UltraGrid3.Rows(_RowIndex).Cells(4).Appearance.BackColor = Color.LightGreen
                    '    UltraGrid3.Rows(_RowIndex).Cells(5).Appearance.BackColor = Color.LightGreen
                    '    UltraGrid3.Rows(_RowIndex).Cells(6).Appearance.BackColor = Color.LightGreen
                    '    UltraGrid3.Rows(_RowIndex).Cells(7).Appearance.BackColor = Color.LightGreen
                    '    UltraGrid3.Rows(_RowIndex).Cells(8).Appearance.BackColor = Color.LightGreen
                    '    UltraGrid3.Rows(_RowIndex).Cells(9).Appearance.BackColor = Color.LightGreen
                    '    UltraGrid3.Rows(_RowIndex).Cells(10).Appearance.BackColor = Color.LightGreen
                    '    UltraGrid3.Rows(_RowIndex).Cells(11).Appearance.BackColor = Color.LightGreen
                    '    UltraGrid3.Rows(_RowIndex).Cells(12).Appearance.BackColor = Color.LightGreen
                    '    UltraGrid3.Rows(_RowIndex).Cells(13).Appearance.BackColor = Color.LightGreen
                    '    UltraGrid3.Rows(_RowIndex).Cells(14).Appearance.BackColor = Color.LightGreen
                    '    UltraGrid3.Rows(_RowIndex).Cells(15).Appearance.BackColor = Color.LightGreen
                    '    UltraGrid3.Rows(_RowIndex).Cells(16).Appearance.BackColor = Color.LightGreen
                    '    UltraGrid3.Rows(_RowIndex).Cells(17).Appearance.BackColor = Color.LightGreen
                    '    UltraGrid3.Rows(_RowIndex).Cells(18).Appearance.BackColor = Color.LightGreen
                    '    UltraGrid3.Rows(_RowIndex).Cells(19).Appearance.BackColor = Color.LightGreen
                    '    UltraGrid3.Rows(_RowIndex).Cells(20).Appearance.BackColor = Color.LightGreen

                    '    ElseIf Trim(UltraGrid3.Rows(_RowIndex).Cells(20).Text) <> "" Then

                    '    UltraGrid3.Rows(_RowIndex).Cells(0).Appearance.BackColor = Color.LightGreen
                    '    UltraGrid3.Rows(_RowIndex).Cells(1).Appearance.BackColor = Color.LightGreen
                    '    UltraGrid3.Rows(_RowIndex).Cells(2).Appearance.BackColor = Color.LightGreen
                    '    UltraGrid3.Rows(_RowIndex).Cells(3).Appearance.BackColor = Color.LightGreen
                    '    UltraGrid3.Rows(_RowIndex).Cells(4).Appearance.BackColor = Color.LightGreen
                    '    UltraGrid3.Rows(_RowIndex).Cells(5).Appearance.BackColor = Color.LightGreen
                    '    UltraGrid3.Rows(_RowIndex).Cells(6).Appearance.BackColor = Color.LightGreen
                    '    UltraGrid3.Rows(_RowIndex).Cells(7).Appearance.BackColor = Color.LightGreen
                    '    UltraGrid3.Rows(_RowIndex).Cells(8).Appearance.BackColor = Color.LightGreen
                    '    UltraGrid3.Rows(_RowIndex).Cells(9).Appearance.BackColor = Color.LightGreen
                    '    UltraGrid3.Rows(_RowIndex).Cells(10).Appearance.BackColor = Color.LightGreen
                    '    UltraGrid3.Rows(_RowIndex).Cells(11).Appearance.BackColor = Color.LightGreen
                    '    UltraGrid3.Rows(_RowIndex).Cells(12).Appearance.BackColor = Color.LightGreen
                    '    UltraGrid3.Rows(_RowIndex).Cells(13).Appearance.BackColor = Color.LightGreen
                    '    UltraGrid3.Rows(_RowIndex).Cells(14).Appearance.BackColor = Color.LightGreen
                    '    UltraGrid3.Rows(_RowIndex).Cells(15).Appearance.BackColor = Color.LightGreen
                    '    UltraGrid3.Rows(_RowIndex).Cells(16).Appearance.BackColor = Color.LightGreen
                    '    UltraGrid3.Rows(_RowIndex).Cells(17).Appearance.BackColor = Color.LightGreen
                    '    UltraGrid3.Rows(_RowIndex).Cells(18).Appearance.BackColor = Color.LightGreen
                    '    UltraGrid3.Rows(_RowIndex).Cells(19).Appearance.BackColor = Color.LightGreen
                    '    UltraGrid3.Rows(_RowIndex).Cells(20).Appearance.BackColor = Color.LightGreen
                    'End If
                End If
                '==========================================================================================

                _RowIndex = _RowIndex + 1
                I = I + 1
            Next

            Dim aDate1 As DateTime
            Dim _Formdate2 As Date
            Dim _Lastdate1 As Date
            Dim L_DateofMonth1 As Integer
            Dim EndDate1 As Date

            aDate1 = Month(txtFromDate.Text) & "/1/" & Year(txtFromDate.Text)
            ' _Formdate1 = aDate
            EndDate1 = aDate1.AddDays(DateTime.DaysInMonth(aDate1.Year, aDate1.Month) - 1)
            '  MsgBox(MonthName(T01.Tables(0).Rows(i)("T01month"), True))
            L_DateofMonth1 = Microsoft.VisualBasic.Day(EndDate1)
            _Lastdate1 = Month(txtFromDate.Text) & "/" & L_DateofMonth1 & "/" & Year(txtFromDate.Text)

            I = 0
            If cboMerch.Text <> "" Then
                Sql = "select F.M08Sales_Order, F.M08Line_Item, sum(F.QTY) AS Qty, MAX(F.Material) AS Material,max(F.Dis) as Dis from View_DelivaryForcus S right OUTER JOIN  View_FG_StockComment F " & _
                    " on S.M07Sales_Order=f.M08Sales_Order and S.M07Line_Item=f.M08Line_Item " & _
                    " where d_date between '" & aDate1 & "' and '" & _Lastdate1 & "' and S.M07Sales_Order is null and  f.M08Merchant='" & cboMerch.Text & "' GROUP BY F.M08Sales_Order, F.M08Line_Item "
            Else
                Sql = "select F.M08Sales_Order, F.M08Line_Item, sum(F.QTY) AS Qty, MAX(F.Material) AS Material,max(F.Dis) as Dis from View_DelivaryForcus S right OUTER JOIN  View_FG_StockComment F " & _
                      " on S.M07Sales_Order=f.M08Sales_Order and S.M07Line_Item=f.M08Line_Item " & _
                      " where d_date between '" & aDate1 & "' and '" & _Lastdate1 & "' and S.M07Sales_Order is null GROUP BY F.M08Sales_Order, F.M08Line_Item "
            End If
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)

            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                Dim newRow3 As DataRow = c_dataCustomer.NewRow
                Dim Value As Double
                Dim _Dis As String

                'If DBNull.Equals(M01.Tables(0).Rows(I)("M08Sales_Order")) = True Then
                'Else
                newRow3("Sales Order") = M01.Tables(0).Rows(I)("M08Sales_Order")
                newRow3("Line Item") = M01.Tables(0).Rows(I)("M08Line_Item")
                ' End If

                newRow3("Material") = M01.Tables(0).Rows(I)("Material")
                newRow3("Material Description") = M01.Tables(0).Rows(I)("Dis")

                Value = M01.Tables(0).Rows(I)("Qty")
                _TotalDel_Qty = _TotalDel_Qty + Value
                _Dis = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _Dis = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0}", Value))
                newRow3("Delivary from Stock") = _Dis

                c_dataCustomer.Rows.Add(newRow3)

                UltraGrid3.Rows(_RowIndex).Cells(0).Appearance.BackColor = Color.LightPink
                UltraGrid3.Rows(_RowIndex).Cells(1).Appearance.BackColor = Color.LightPink
                UltraGrid3.Rows(_RowIndex).Cells(2).Appearance.BackColor = Color.LightPink
                UltraGrid3.Rows(_RowIndex).Cells(3).Appearance.BackColor = Color.LightPink
                UltraGrid3.Rows(_RowIndex).Cells(4).Appearance.BackColor = Color.LightPink
                UltraGrid3.Rows(_RowIndex).Cells(5).Appearance.BackColor = Color.LightPink
                UltraGrid3.Rows(_RowIndex).Cells(6).Appearance.BackColor = Color.LightPink
                UltraGrid3.Rows(_RowIndex).Cells(7).Appearance.BackColor = Color.LightPink
                UltraGrid3.Rows(_RowIndex).Cells(8).Appearance.BackColor = Color.LightPink
                UltraGrid3.Rows(_RowIndex).Cells(9).Appearance.BackColor = Color.LightPink
                UltraGrid3.Rows(_RowIndex).Cells(10).Appearance.BackColor = Color.LightPink
                UltraGrid3.Rows(_RowIndex).Cells(11).Appearance.BackColor = Color.LightPink
                UltraGrid3.Rows(_RowIndex).Cells(12).Appearance.BackColor = Color.LightPink
                UltraGrid3.Rows(_RowIndex).Cells(13).Appearance.BackColor = Color.LightPink
                UltraGrid3.Rows(_RowIndex).Cells(14).Appearance.BackColor = Color.LightPink
                UltraGrid3.Rows(_RowIndex).Cells(15).Appearance.BackColor = Color.LightPink
                UltraGrid3.Rows(_RowIndex).Cells(16).Appearance.BackColor = Color.LightPink

                _RowIndex = _RowIndex + 1
                I = I + 1
            Next
            Dim Value1 As Double
            Dim _Dis1 As String
            Dim _Index As Integer
            Dim _ProQ1 As Double
            Dim _ProQ2 As Double
            Dim _ProQ3 As Double

            _ProQ1 = 0
            _ProQ2 = 0
            _ProQ3 = 0

            For Each uRow As UltraGridRow In UltraGrid3.Rows
                With UltraGrid3
                    If IsNumeric(.Rows(_Index).Cells(9).Value) Then
                        _ProQ1 = _ProQ1 + .Rows(_Index).Cells(9).Value

                    End If
                    If IsNumeric(.Rows(_Index).Cells(10).Value) Then
                        _ProQ2 = _ProQ2 + .Rows(_Index).Cells(10).Value

                    End If

                    'If IsNumeric(.Rows(_Index).Cells(9).Value) Then
                    '    _ProQ3 = _ProQ3 + .Rows(_Index).Cells(7).Value

                    'End If

                End With

                _Index = _Index + 1
            Next

            Dim newRow1 As DataRow = c_dataCustomer.NewRow
            Value1 = _DelQty

            _Dis1 = (Value1.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _Dis1 = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0}", Value1))
            newRow1("Possi Qty From WIP") = _Dis1

            Value1 = _BTQty01
            If _BT1 <> "" Then
                _Dis1 = (Value1.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _Dis1 = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0}", Value1))
                newRow1(_BT1) = _Dis1
            End If
            Value1 = _BTQty02
            If _BT2 <> "" Then
                _Dis1 = (Value1.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _Dis1 = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0}", Value1))
                newRow1(_BT2) = _Dis1
            End If
            Value1 = _BTQty03

            'If _BT3 <> "" Then
            '    _Dis1 = (Value1.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            '    _Dis1 = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0}", Value1))
            '    newRow1(_BT3) = _Dis1
            'End If
            Value1 = _ProQ1
            _Dis1 = (Value1.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _Dis1 = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0}", Value1))
            newRow1(_ProDis1) = _Dis1

            Value1 = _ProQ2
            _Dis1 = (Value1.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _Dis1 = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0}", Value1))
            newRow1(_ProDis2) = _Dis1

            Value1 = _TobeWkQty
            _Dis1 = (Value1.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _Dis1 = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0}", Value1))
            newRow1(_TobeDis) = _Dis1


            Value1 = _TobeWKQty1
            _Dis1 = (Value1.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _Dis1 = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0}", Value1))
            newRow1(_TobeDis1) = _Dis1

            'Value1 = _Stok1
            '_Dis1 = (Value1.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            '_Dis1 = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0}", Value1))
            'newRow1(_StockDis1) = _Dis1

            'Value1 = _Stock2
            '_Dis1 = (Value1.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            '_Dis1 = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0}", Value1))
            'newRow1(_StockDis2) = _Dis1

            newRow1("WIP(Exam)") = _WIPExam
            newRow1("WIP(Finishing)") = _WIPFinishing
            newRow1("WIP(Dyeing)") = _WIPDye
            Value1 = _Total_PossibleQTY1
            ' _Total_PossibleQTY = _Total_PossibleQTY + Value
            _Dis1 = (Value1.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _Dis1 = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0}", Value1))
            newRow1("Possi Qty From WIP") = _Dis1

            Value1 = _TotalDel_Qty
            _Dis1 = (Value1.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _Dis1 = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0}", Value1))
            newRow1("Delivary from Stock") = _Dis1


            Value1 = _DelQty
            _Dis1 = (Value1.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _Dis1 = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0}", Value1))
            newRow1("Del.qty as today") = _Dis1
            ' _RowIndex = _RowIndex + 1
            c_dataCustomer.Rows.Add(newRow1)

            UltraGrid3.Rows(_RowIndex).Cells(4).Appearance.BackColor = Color.DeepSkyBlue
            UltraGrid3.Rows(_RowIndex).Cells(5).Appearance.BackColor = Color.DeepSkyBlue
            UltraGrid3.Rows(_RowIndex).Cells(6).Appearance.BackColor = Color.DeepSkyBlue
            UltraGrid3.Rows(_RowIndex).Cells(7).Appearance.BackColor = Color.DeepSkyBlue
            UltraGrid3.Rows(_RowIndex).Cells(8).Appearance.BackColor = Color.DeepSkyBlue
            UltraGrid3.Rows(_RowIndex).Cells(9).Appearance.BackColor = Color.DeepSkyBlue
            UltraGrid3.Rows(_RowIndex).Cells(10).Appearance.BackColor = Color.DeepSkyBlue
            UltraGrid3.Rows(_RowIndex).Cells(11).Appearance.BackColor = Color.DeepSkyBlue
            UltraGrid3.Rows(_RowIndex).Cells(12).Appearance.BackColor = Color.DeepSkyBlue
            UltraGrid3.Rows(_RowIndex).Cells(13).Appearance.BackColor = Color.DeepSkyBlue
            UltraGrid3.Rows(_RowIndex).Cells(14).Appearance.BackColor = Color.DeepSkyBlue
            UltraGrid3.Rows(_RowIndex).Cells(15).Appearance.BackColor = Color.DeepSkyBlue
            UltraGrid3.Rows(_RowIndex).Cells(16).Appearance.BackColor = Color.DeepSkyBlue
            '  UltraGrid3.Rows(_RowIndex).Cells(17).Appearance.BackColor = Color.DeepSkyBlue
            ' UltraGrid3.Rows(_RowIndex).Cells(18).Appearance.BackColor = Color.DeepSkyBlue
            'UltraGrid3.Rows(_RowIndex).Cells(19).Appearance.BackColor = Color.DeepSkyBlue
            'UltraGrid3.Rows(_RowIndex).Cells(20).Appearance.BackColor = Color.DeepSkyBlue
            '   UltraGrid3.Rows(_RowIndex).Cells(20).Appearance.BackColor = Color.DeepSkyBlue
            ' MsgBox(UltraGrid3.Rows.Count)
            'I = 0
            'Dim Z As Double
            'For Each uRow As UltraGridRow In UltraGrid3.Rows
            '    If UltraGrid3.Rows(I).Cells(5).Text <> "" Then
            '        Z = Z + CDbl(UltraGrid3.Rows(I).Cells(5).Value)
            '    End If
            '    I = I + 1
            'Next

            con.Close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                MsgBox(I)
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
                con.Close()
            End If
        End Try

    End Function

    Private Sub cboMaterial_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboMaterial.AfterCloseUp
        ' Call Find_Data()
    End Sub

    Private Sub cboMaterial_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles cboMaterial.InitializeLayout

    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        Call Load_Gride_SalesOrder()
        Call Find_Data()
    End Sub

    Private Sub UltraGrid3_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles UltraGrid3.AfterCellUpdate
        'Dim _Index As Integer
        '_Index = UltraGrid3.Rows.Count
        'If _Index > 0 Then
        '    UltraGrid3.Rows(_Index - 1).Cells(8).Value = _ProQ1
        '    UltraGrid3.Rows(_Index - 1).Cells(9).Value = _ProQ2
        'End If
    End Sub

    Private Sub UltraGrid3_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles UltraGrid3.KeyUp

        Try
            If e.KeyCode = Keys.Escape Then
                Dim _Index As Integer
                Dim _LastRow As Integer

                _LastRow = UltraGrid3.Rows.Count
                _ProQ1 = 0
                _ProQ2 = 0
                _ProQ3 = 0
                _Index = 0
                For Each uRow As UltraGridRow In UltraGrid3.Rows
                    If _Index = UltraGrid3.Rows.Count - 1 Then

                    Else
                        With UltraGrid3
                            If IsNumeric(.Rows(_Index).Cells(9).Value) Then
                                _ProQ1 = _ProQ1 + .Rows(_Index).Cells(9).Value
                                ' .Rows(_LastRow - 1).Cells(8).Value = _ProQ1
                            End If
                            If IsNumeric(.Rows(_Index).Cells(10).Value) Then
                                _ProQ2 = _ProQ2 + .Rows(_Index).Cells(10).Value

                            End If

                            'If IsNumeric(.Rows(_Index).Cells(9).Value) Then
                            '    _ProQ3 = _ProQ3 + .Rows(_Index).Cells(7).Value

                            'End If

                        End With
                    End If
                    _Index = _Index + 1
                Next

                _Index = UltraGrid3.Rows.Count
                If _Index > 0 Then
                    UltraGrid3.Rows(_Index - 1).Cells(9).Value = _ProQ1
                    UltraGrid3.Rows(_Index - 1).Cells(10).Value = _ProQ2
                End If
            End If

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                '  DBEngin.CloseConnection(connection)
                ' connection.ConnectionString = ""
            End If
        End Try
    End Sub



    Private Sub UltraGrid3_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles UltraGrid3.TextChanged
        Dim _Index As Integer
        ' _Index = e.row.index
        _Index = UltraGrid3.Rows.Band.Index
        With UltraGrid3

            ' .Rows(_Index).Cells(11).Value()
        End With
    End Sub

    Private Sub cboMaterial_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboMaterial.TextChanged

    End Sub

    Private Sub chkCus_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkCus.CheckedChanged

    End Sub

    Private Sub UltraButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton1.Click
        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean
        Dim M01 As DataSet

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Dim i As Integer
        Dim _Fromdate As Date
        Dim _Todate As Date
        Dim _Pr1 As Double
        Dim _Pr2 As Double
        Dim _Value As Double

        Try
            _Fromdate = txtFromDate.Text

            _Fromdate = _Fromdate
            Dim thisCulture = Globalization.CultureInfo.CurrentCulture
            Dim dayOfWeek As DayOfWeek = thisCulture.Calendar.GetDayOfWeek(_Fromdate)
            Dim dayName As String = thisCulture.DateTimeFormat.GetDayName(dayOfWeek)

            Dim daysInFeb As Integer = System.DateTime.DaysInMonth(Year(_Fromdate), Month(_Fromdate))
            _Todate = CDate(_Fromdate).AddDays(+daysInFeb)
            _Todate = _Todate.AddDays(-1)
            i = 0
            ' MsgBox(UltraGrid3.Rows.Count)
            For Each uRow As UltraGridRow In UltraGrid3.Rows
                'QUARANTINE REASON FOR REPORT TABLE
                _Pr1 = 0
                _Pr2 = 0
                If i = 13 Then
                    ' MsgBox("")
                End If
                With UltraGrid3

                    If UltraGrid3.Rows.Count - 1 = i Then
                        Exit For
                    End If
                    If Trim(.Rows(i).Cells(10).Text) <> "" Then
                        _Pr1 = .Rows(i).Cells(10).Value
                    End If
                    If Trim(.Rows(i).Cells(11).Text) <> "" Then
                        _Pr2 = .Rows(i).Cells(11).Value
                    End If

                    Dim _ConFact As Double
                    Dim _ValueKG As Double

                    _Value = 0
                    _ConFact = 0
                    _ValueKG = 0

                    nvcFieldList1 = "select * from M07TobeDelivered where M07Sales_Order='" & Trim(.Rows(i).Cells(0).Value) & "' and M07Line_Item='" & Trim(.Rows(i).Cells(1).Value) & "'"
                    dsUser = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                    If isValidDataset(dsUser) Then
                        _Value = dsUser.Tables(0).Rows(0)("M07Unit_PriceMTR")
                        _Merchant = dsUser.Tables(0).Rows(0)("M07Merchant")
                        If dsUser.Tables(0).Rows(0)("M07Qty_Kg") > 0 Then
                            _ConFact = dsUser.Tables(0).Rows(0)("M07Qty_Mtr") / dsUser.Tables(0).Rows(0)("M07Qty_Kg")
                        End If
                        _ValueKG = dsUser.Tables(0).Rows(0)("M07Unit_PriceKG")
                    End If
                    If _Pr1 > 0 Or _Pr2 > 0 Then
                        nvcFieldList1 = "select * from T08DelayComment where T08Sales_Order='" & Trim(.Rows(i).Cells(0).Value) & "' and T08Line_Item='" & Trim(.Rows(i).Cells(1).Value) & "' and T08Material='" & Trim(.Rows(i).Cells(2).Value) & "'  order by T08Date  DESC"
                        M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                        If isValidDataset(M01) Then
                            nvcFieldList1 = "update T08DelayComment set T08Date='" & Today & "',T08Pro_Qty1='" & _Pr1 & "',T08Pro_Qty2='" & _Pr2 & "',T08Value='" & _Value & "',T08Merchant='" & _Merchant & "',T08Con_Fact='" & _ConFact & "',T08Value_KG='" & _ValueKG & "' where T08Sales_Order='" & Trim(.Rows(i).Cells(0).Value) & "' and T08Line_Item='" & Trim(.Rows(i).Cells(1).Value) & "' and T08Material='" & Trim(.Rows(i).Cells(2).Value) & "'  "
                            ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                        Else
                            nvcFieldList1 = "Insert Into T08DelayComment(T08Sales_Order,T08Line_Item,T08Material,T08Date,T08Pro_Qty1,T08Pro_Qty2,T08Value,T08Merchant,T08Value_KG,T08Con_Fact)" & _
                                                   " values('" & .Rows(i).Cells(0).Value & "','" & .Rows(i).Cells(1).Value & "','" & .Rows(i).Cells(2).Value & "','" & Today & "','" & _Pr1 & "','" & _Pr2 & "','" & _Value & "','" & _Merchant & "','" & _ValueKG & "','" & _ConFact & "')"
                            ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                        End If
                    End If
                End With

                i = i + 1
            Next
            MsgBox("Records successfully updated", MsgBoxStyle.Information, "Information ..........")
            transaction.Commit()
            connection.Close()
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
            common.ClearAll(OPR0)
            OPR0.Enabled = True
            ' OPR2.Enabled = True
            Call Load_Gride_SalesOrder()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                MsgBox(i)
                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
                connection.Close()
            End If
        End Try

    End Sub

    Private Sub UltraCheckEditor1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkUpload.CheckedChanged
        If chkUpload.Checked = True Then
            Call Upload_M09ZPL_ORDER()
            Call Upload_M08Stock()
            Call Upload_TobeDeliverd_File()
            Call Upload_Delivary_File()
            Call FR_Update()
            MsgBox("Record uploaded successfully", MsgBoxStyle.Information, "Information ....")

        End If
    End Sub

    Function FR_Update()
        Dim sr As System.IO.StreamReader
        Dim strFileName As String
        Dim _PO_No As String
        Dim _sales_Order As String
        Dim _LineItem As String
        Dim _Shadule As String
        Dim _Material As String
        Dim _Material_Dis As String
        Dim _Customer As String
        Dim _Department As String
        Dim _Merchnat As String
        Dim _Del_Date As String
        Dim _Order_Qty As Double
        Dim _Del_Qty As Double
        Dim _Delay_Qty As Double
        Dim _FGStock As Double
        Dim _Balance As Double
        Dim _Location As String
        Dim _PRD_Qty As String
        Dim _Grg_Qty As Double
        Dim _NCComment As String
        Dim _Awaiting As String
        Dim _depComm As String
        Dim _Comm2 As String
        Dim _OTDStatus As String
        Dim _PRD_OrderQty As Double
        Dim _AppStatus As String


        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Dim M01 As DataSet
        Dim I As Integer
        Dim A As String
        Dim characterToRemove As String
        Dim _DyeMC As String

        Dim X11 As Integer
        Dim _CusCode As String

        Try
            nvcFieldList1 = "delete from FR_Update"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            strFileName = ConfigurationManager.AppSettings("FilePath") + "\FR_otd.txt"
            For Each line As String In System.IO.File.ReadAllLines(strFileName)
                Dim fields() As String = line.Split(vbTab)
                If X11 = 278 Then
                    '  MsgBox("")
                End If
                ' _Location = Trim(fields(15))
                'If _Location <> "" Then

                If X11 = 0 Then

                Else
                    _sales_Order = Trim(fields(0))
                    _LineItem = Trim(fields(1))

                    _Department = Trim(fields(2))
                    _Del_Date = Trim(fields(3))

                    _DyeMC = Trim(fields(4))
                    _AppStatus = Trim(fields(5))
                    '_OTDStatus = "Fales"

                    characterToRemove = "-"
                    'MsgBox(Trim(fields(9)))
                    _Del_Date = (Replace(_Del_Date, characterToRemove, "/"))
                    Dim A1 As String

                    ' A1 = (Microsoft.VisualBasic.Left(_Del_Date, 5))
                    '_Del_Date = Microsoft.VisualBasic.Right(A1, 2) & "/" & Microsoft.VisualBasic.Left(A1, 2) & "/" & Microsoft.VisualBasic.Right(_Del_Date, 4)
                    'Dim oDate As DateTime = Convert.ToDateTime(_Del_Date)
                    'MsgBox(oDate.Day & " " & oDate.Month & "  " & oDate.Year)



                    nvcFieldList1 = "SELECT * FROM FR_Update WHERE Batch_No='" & Trim(_sales_Order) & "'"
                    M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                    If isValidDataset(M01) Then
                        nvcFieldList1 = "update FR_Update set Stock_Code='" & _LineItem & "',Recipy_Status='" & _Department & "',Dye_Pln_Date='" & _Del_Date & "',Dye_Machine='" & _DyeMC & "',first_BlkApp='" & _AppStatus & "' where Batch_No='" & _sales_Order & "'"
                        ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                    Else
                        nvcFieldList1 = "Insert Into FR_Update(Batch_No,Stock_Code,Recipy_Status,Dye_Pln_Date,Dye_Machine,first_BlkApp)" & _
                                                            " values('" & Trim(_sales_Order) & "', '" & Trim(_LineItem) & "','" & Trim(_Department) & "','" & Trim(_Del_Date) & "','" & _DyeMC & "','" & _AppStatus & "')"
                        ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                    End If
                    _CusCode = ""
                    _Customer = ""
                    _sales_Order = ""
                    _Department = ""
                    _LineItem = ""
                    _OTDStatus = ""

                    'cmdEdit.Enabled = True
                End If
                X11 = X11 + 1
            Next
            'MsgBox("Record Update successfully", MsgBoxStyle.Information, "Information ....")
            transaction.Commit()
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
        Catch malFormLineEx As Microsoft.VisualBasic.FileIO.MalformedLineException
            MessageBox.Show("Line " & malFormLineEx.Message & "is not valid and will be skipped.", "Malformed Line Exception")
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
        Catch ex As Exception
            MessageBox.Show(ex.Message & " exception has occurred.", "Exception")
            MsgBox(X11 & "-" & strFileName)
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
        Finally
            ' If successful or if an exception is thrown,
            ' close the TextFieldParser.
            ' theTextFieldParser.Close()
        End Try
    End Function

    Function Upload_M09ZPL_ORDER()
        Dim strFileName As String
        strFileName = "\\Tjlapp04\grginspec_dload$\TJL CAT.txt"
        strFileName = "E:\TJL_MILAN\SAP_DOWNLOADS\Sales Forcust\ZPL_ORDER.txt"
        'Dim strFileName As String
        strFileName = ConfigurationManager.AppSettings("FilePath") + "\ZPL_ORDER.txt"
        Dim CurrGameWinningSerials(0 To MAX_SERIALS) As Long
        Dim fileHndl As Long
        Dim lLineNo As Long

        Dim strOrder, strLineItem, _
      strMerchant, strDis As String
        Dim strDep As String
        Dim strSpec As Double

        Dim strKg As Double
        Dim strMtr As Double
        ' Dim strFileName As String '= _
        Dim strDate As String
        ' Dim strDate As String
        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True

        Dim M03Knittingorder As DataSet
        Dim ncQryType As String
        Dim nvcVccode As String
        Dim linesList As New List(Of String)(IO.File.ReadAllLines(strFileName))
        Dim _RefNo As Integer
        Dim P01Parameter As DataSet
        Dim _Value As Double
        Dim M06Cls As DataSet
        Dim str30Class As String
        Dim strMaterial As String

        Dim nvcFieldList As String
        Dim nvcWhereClause As String
        Dim nvcVcLine As String
        Dim nvcVcDate As String
        Dim nvcBatch As String
        Dim nvcQtype As String
        Dim strBatch As String
        Dim strCustomer As String
        Dim strLast_Con_Date As String
        Dim strNext_Opp As String
        Dim strPlanning_Comm As String
        Dim strOrder_Type As String
        Dim strRetailer As String
        Dim strFinisg_MC As String
        Dim strTBC As String
        Dim strCon_Fact As String
        Dim strFabric_Type As String
        Dim _ZPLType As String

        Dim characterToRemove As String


        Try

            nvcFieldList1 = "delete from M09ZPL_ORDER "
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            fileHndl = FreeFile()


            ' strFileName = Dir(strFileName)

            'UPGRADE_WARNING: Couldn't resolve default property of object fileHndl. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object strValidSerialFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            FileOpen(fileHndl, strFileName, OpenMode.Input)
            lLineNo = 0
            Dim strRow As String

            Do Until EOF(fileHndl)


                '  Line Input #fileHndl, strRow
                'UPGRADE_WARNING: Couldn't resolve default property of object fileHndl. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                strRow = LineInput(fileHndl)

                If lLineNo = 1124 Then
                    ' MsgBox("")
                End If
                If Trim(strRow) <> "" Then
                    ncQryType = "LST"
                    If InStr(1, strRow, vbTab) > 0 Then
                        If (Trim(Split(strRow, vbTab)(11))) <> "" Then
                            '  CurrGameWinningSerials(lLineNo) = Trim(Split(strRow, vbTab)(0))
                            strBatch = (Trim(Split(strRow, vbTab)(0)))
                            strCustomer = (Trim(Split(strRow, vbTab)(1)))
                            strMaterial = CInt(Trim(Split(strRow, vbTab)(2)))
                            '    strMaterial = Microsoft.VisualBasic.Left(strMaterial, 2) & "-" & Microsoft.VisualBasic.Right(strMaterial, 5)
                            strDis = (Trim(Split(strRow, vbTab)(3)))

                            strDate = (Trim(Split(strRow, vbTab)(4)))
                            strDate = Microsoft.VisualBasic.Right(Microsoft.VisualBasic.Left(strDate, 6), 2)
                            strDate = strDate & "/" & Microsoft.VisualBasic.Right(Trim(Split(strRow, vbTab)(4)), 2)
                            strDate = strDate & "/" & Microsoft.VisualBasic.Left(Trim(Split(strRow, vbTab)(4)), 4)
                            If (Trim(Split(strRow, vbTab)(5))) = "0" Then
                            Else
                                If (Trim(Split(strRow, vbTab)(5))) <> "" Then
                                    strLast_Con_Date = (Trim(Split(strRow, vbTab)(5)))
                                    strLast_Con_Date = Microsoft.VisualBasic.Right(Microsoft.VisualBasic.Left(strLast_Con_Date, 6), 2)
                                    strLast_Con_Date = strLast_Con_Date & "/" & Microsoft.VisualBasic.Right(Trim(Split(strRow, vbTab)(5)), 2)
                                    strLast_Con_Date = strLast_Con_Date & "/" & Microsoft.VisualBasic.Left(Trim(Split(strRow, vbTab)(5)), 4)
                                End If

                            End If

                            strNext_Opp = (Trim(Split(strRow, vbTab)(7)))
                            strPlanning_Comm = (Trim(Split(strRow, vbTab)(8)))
                            strOrder_Type = (Trim(Split(strRow, vbTab)(9)))
                            If (Trim(Split(strRow, vbTab)(10))) = "ZP12" Or (Trim(Split(strRow, vbTab)(10))) = "ZP14" Then

                            Else
                                strBatch = CInt(strBatch)
                            End If

                            _ZPLType = (Trim(Split(strRow, vbTab)(10)))
                            strOrder = CInt(Trim(Split(strRow, vbTab)(11)))
                            strLineItem = CInt(Trim(Split(strRow, vbTab)(12)))
                            strMerchant = (Trim(Split(strRow, vbTab)(14)))
                            strCon_Fact = (Trim(Split(strRow, vbTab)(17)))
                            strFabric_Type = (Trim(Split(strRow, vbTab)(18)))
                            strFinisg_MC = (Trim(Split(strRow, vbTab)(19)))
                            strTBC = (Trim(Split(strRow, vbTab)(20)))

                            characterToRemove = "'"
                            strTBC = (Replace(strTBC, characterToRemove, ""))

                            strDis = (Replace(strDis, characterToRemove, ""))
                            strCustomer = (Replace(strCustomer, characterToRemove, ""))
                            strPlanning_Comm = (Replace(strPlanning_Comm, characterToRemove, ""))
                            M06Cls = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetM09ZPL_ORDER", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcCode", strOrder), New SqlParameter("@vcLine", strLineItem), New SqlParameter("@vcDate", strDate), New SqlParameter("@vcBATCH", strBatch))
                            If isValidDataset(M06Cls) Then
                                For Each DTRow As DataRow In M06Cls.Tables(0).Rows
                                    nvcWhereClause = "M09Sales_Oredr='" & strOrder & "' AND M09Line_Item='" & strLineItem & "' AND M09Del_Date='" & strDate & "' AND M09BatchNo='" & nvcBatch & "'"
                                    ncQryType = "UPD"
                                    nvcFieldList1 = "M09Qty_KG='" & (Trim(Split(strRow, vbTab)(6))) & "',M09Qty_Mtr='" & (Trim(Split(strRow, vbTab)(13))) & "'"
                                    up_GetSetM09ZPL_ORDER(ncQryType, nvcFieldList1, nvcWhereClause, nvcVcLine, nvcVcDate, nvcBatch, nvcVccode, connection, transaction)
                                    ' ExecuteNonQueryText(connection, transaction, "up_GetSetM06Delivary_Qty", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcCode", strOrder), New SqlParameter("@vcLine", strLineItem), New SqlParameter("@vcDate", strDate))
                                Next
                            Else
                                ncQryType = "ADD"
                                nvcFieldList1 = "(M09BatchNo," & "M09Customer," & "M09Meterial," & "M09Dis," & "M09Del_Date," & "M09Lat_Con_Date," & "M09Qty_KG," & "M09Next_Opp," & "M09Planning_Comm," & "M09Oredr_Type," & "M09Sales_Oredr," & "M09Line_Item," & "M09Qty_Mtr," & "M09Merchant," & "M09Con_Fact," & "M09Fabric_Type," & "M09Finishing_MC," & "M09TBC," & "M09ZPL_OrderType) " & "values('" & Trim(strBatch) & "','" & strCustomer & "','" & strMaterial & "','" & strDis & "','" & strDate & "','" & strLast_Con_Date & "','" & (Trim(Split(strRow, vbTab)(6))) & "','" & strNext_Opp & "','" & strPlanning_Comm & "','" & strOrder_Type & "','" & strOrder & "','" & strLineItem & "','" & (Trim(Split(strRow, vbTab)(13))) & "','" & strMerchant & "','" & strCon_Fact & "','" & strFabric_Type & "','" & strFinisg_MC & "','" & strTBC & "','" & _ZPLType & "')"
                                up_GetSetM09ZPL_ORDER(ncQryType, nvcFieldList1, nvcWhereClause, nvcVcLine, nvcVcDate, nvcBatch, nvcVccode, connection, transaction)
                            End If
                            '---------------------------------------------------------------------------------------------
                            '        nvcFieldList1 = "select * from M02Defect where M02OrderNo='" & strOrder & "' and M02Roll='" & strCATRoll & "' and M02SeqNo='" & strSqNo & "'"
                            '        M03Knittingorder = DbEngine.ExecuteDataset(connection, transaction, nvcFieldList1)
                            '        If isValidDataset(M03Knittingorder) Then
                            '            ' nvcFieldList1 = "UPDATE M04Cutoff set M03Orderqty=" & strOrderqty & ",M03Yarnstock='" & strYarncode & "',M03YarnType='" & strYarntype & "',M03IType='" & strI_Type & "',M03LineItem='" & strLineItem & "',M0330Class='" & str30class & "',M03Root='" & strRoot & "',M03MCNo='" & strMC & "',M03CuttingLine='" & strCutting_line & "',M03NoofRoll='" & strRolls & "' where M03OrderNo='" & strOrder & "' and M03Quality='" & strQuality & "' and M03Material='" & strMaterial & "'"
                            '            'ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                            '        Else
                            '            ncQryType = "ADD"
                            '            nvcFieldList1 = "(M02RefNo," & "M02OrderNo," & "M02Roll," & "M02SeqNo," & "M02Code," & "M02Spec," & "M02Dep," & "M02Date," & "M02Status," & "M02Class," & "M02Reason) " & "values(" & _RefNo & ",'" & strOrder & "'," & strCATRoll & ",'" & strSqNo & "','" & strCode & "'," & strSpec & ",'" & strDep & "','" & strDate & "','I','" & str30Class & "','" & strCode & "')"
                            '            'nvcFieldList1 = "(M04Order," & "M04CATRoll," & "M04SAPRoll) " & "values('" & Trim(strOrder) & "','" & strCATRoll & "','" & strSAPRoll & "')"
                            '            up_GetSetM02Defect(ncQryType, nvcFieldList1, nvcVccode, connection, transaction)


                            '            '---------------------------------------------------UPDATE P01PARAMETER TABLE
                            '            nvcFieldList1 = "UPDATE P01Parameter SET P01NO=P01NO +" & 1 & " WHERE P01CODE='DF'"
                            '            ExecuteNonQueryText(connection, transaction, nvcFieldList1)



                            '        End If
                            '    End If
                            'End If
                            linesList.RemoveAt(0)
                            ''  MsgBox(linesList.ToArray().ToString)
                            'IO.File.WriteAllLines(strFileName, linesList.ToArray())

                            strMtr = 0
                            strKg = 0

                        Else
                            '  Err.Raise(vbObjectError + 18001, "GenerateInstantFile(str,str,str)", "Invalid Record At Line " & CStr(lLineNo))
                        End If
                    End If
                End If

                lLineNo = lLineNo + 1

            Loop

            transaction.Commit()
            DBEngin.CloseConnection(connection)
            ' transaction.Rollback()
            ' MsgBox("M09ZPL_ORDER ")
            connection.Close()
            FileClose()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                MsgBox(lLineNo & "-" & strFileName)
            End If
        End Try
    End Function

    Function Upload_M08Stock()
        Dim strFileName As String
        strFileName = "\\Tjlapp04\grginspec_dload$\TJL CAT.txt"
        ' strFileName = "E:\TJL_MILAN\SAP_DOWNLOADS\Sales Forcust\Stock.txt"
        'Dim strFileName As String
        strFileName = ConfigurationManager.AppSettings("FilePath") + "\Stock_FG.txt"
        Dim CurrGameWinningSerials(0 To MAX_SERIALS) As Long
        Dim fileHndl As Long
        Dim lLineNo As Long

        Dim strOrder, strLineItem, _
      strMerchant, strDis As String
        Dim strDep As String
        Dim strSpec As Double

        Dim strKg As Double
        Dim strMtr As Double
        ' Dim strFileName As String '= _
        Dim strDate As String
        ' Dim strDate As String
        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True

        Dim M03Knittingorder As DataSet
        Dim ncQryType As String
        Dim nvcVccode As String
        Dim linesList As New List(Of String)(IO.File.ReadAllLines(strFileName))
        Dim _RefNo As Integer
        Dim P01Parameter As DataSet
        Dim _Value As Double
        Dim M06Cls As DataSet
        Dim str30Class As String
        Dim strMaterial As String

        Dim nvcFieldList As String
        Dim nvcWhereClause As String
        Dim nvcVcLine As String
        Dim nvcVcDate As String
        Dim nvcQtype As String

        Dim characterToRemove As String


        Try
            nvcFieldList1 = "delete from  M08Stock "
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            fileHndl = FreeFile()


            ' strFileName = Dir(strFileName)

            'UPGRADE_WARNING: Couldn't resolve default property of object fileHndl. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object strValidSerialFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            FileOpen(fileHndl, strFileName, OpenMode.Input)
            lLineNo = 0
            Dim strRow As String

            Do Until EOF(fileHndl)

                '  Line Input #fileHndl, strRow
                'UPGRADE_WARNING: Couldn't resolve default property of object fileHndl. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                strRow = LineInput(fileHndl)
                If lLineNo = 286 Then
                    '   MsgBox("")
                End If

                If Trim(strRow) <> "" Then
                    ncQryType = "LST"
                    If InStr(1, strRow, vbTab) > 0 Then

                        '  CurrGameWinningSerials(lLineNo) = Trim(Split(strRow, vbTab)(0))
                        strOrder = (Trim(Split(strRow, vbTab)(0)))
                        ' strOrder = Microsoft.VisualBasic.Right(strOrder, 7)
                        strLineItem = CInt(Trim(Split(strRow, vbTab)(1)))
                        strDate = (Trim(Split(strRow, vbTab)(9)))
                        strDate = Microsoft.VisualBasic.Right(Microsoft.VisualBasic.Left(strDate, 6), 2)
                        strDate = strDate & "/" & Microsoft.VisualBasic.Right(Trim(Split(strRow, vbTab)(9)), 2)
                        strDate = strDate & "/" & Microsoft.VisualBasic.Left(Trim(Split(strRow, vbTab)(9)), 4)
                        strMaterial = (Trim(Split(strRow, vbTab)(2)))
                        'strMaterial = Microsoft.VisualBasic.Left(strMaterial, 2) & "-" & Microsoft.VisualBasic.Right(strMaterial, 5)
                        strDis = (Trim(Split(strRow, vbTab)(3)))
                        strMerchant = (Trim(Split(strRow, vbTab)(5)))
                        str30Class = (Trim(Split(strRow, vbTab)(4)))
                        characterToRemove = "'"

                        strDis = (Replace(strDis, characterToRemove, ""))
                        characterToRemove = "-"

                        strMaterial = (Replace(strMaterial, characterToRemove, ""))
                        characterToRemove = "'"
                        str30Class = (Replace(str30Class, characterToRemove, ""))
                        ' M06Cls = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetM08Stock", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcCode", strOrder), New SqlParameter("@vcLine", strLineItem)) ', New SqlParameter("@vcDate", strDate))
                        nvcFieldList = "select * from M08Stock where M08Sales_Order='" & strOrder & "' AND M08Line_Item='" & strLineItem & "' and M08RollNo='" & (Trim(Split(strRow, vbTab)(6))) & "'"
                        M06Cls = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList)
                        If isValidDataset(M06Cls) Then
                            For Each DTRow As DataRow In M06Cls.Tables(0).Rows
                                If Trim(Split(strRow, vbTab)(11)) = "9030" Then
                                Else
                                    nvcWhereClause = "M08Sales_Order='" & strOrder & "' AND M08Line_Item='" & strLineItem & "'" ' AND M08TR_Date='" & strDate & "'"
                                    ncQryType = "UPD"
                                    nvcFieldList1 = "M08Qty_Mtr= M08Qty_Mtr +" & (Trim(Split(strRow, vbTab)(11))) & ",M08Qty_KG=M08Qty_KG +" & (Trim(Split(strRow, vbTab)(10))) & ""
                                    up_GetSetM08Stock(ncQryType, nvcFieldList1, nvcWhereClause, nvcVcLine, nvcVcDate, nvcVccode, connection, transaction)
                                    ' ExecuteNonQueryText(connection, transaction, "up_GetSetM06Delivary_Qty", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcCode", strOrder), New SqlParameter("@vcLine", strLineItem), New SqlParameter("@vcDate", strDate))
                                End If
Next
                        Else
                            If Trim(Split(strRow, vbTab)(11)) = "9030" Then
                            Else
                                ncQryType = "ADD"
                                nvcFieldList1 = "(M08Sales_Order," & "M08Line_Item," & "M08TR_Date," & "M08Meterial," & "M08Dis," & "M08Retailer," & "M08Merchant," & "M08RollNo," & "M08Batch_No," & "M08Qty_KG," & "M08Qty_Mtr," & "M08Location) " & "values('" & Trim(strOrder) & "','" & strLineItem & "','" & strDate & "','" & strMaterial & "','" & strDis & "','" & str30Class & "','" & strMerchant & "','" & (Trim(Split(strRow, vbTab)(6))) & "','" & (Trim(Split(strRow, vbTab)(7))) & "','" & (Trim(Split(strRow, vbTab)(10))) & "','" & (Trim(Split(strRow, vbTab)(11))) & "','" & (Trim(Split(strRow, vbTab)(12))) & "' )"
                                up_GetSetM08Stock(ncQryType, nvcFieldList1, nvcWhereClause, nvcVcLine, nvcVcDate, nvcVccode, connection, transaction)
                            End If
                            End If
                            '---------------------------------------------------------------------------------------------
                            '        nvcFieldList1 = "select * from M02Defect where M02OrderNo='" & strOrder & "' and M02Roll='" & strCATRoll & "' and M02SeqNo='" & strSqNo & "'"
                            '        M03Knittingorder = DbEngine.ExecuteDataset(connection, transaction, nvcFieldList1)
                            '        If isValidDataset(M03Knittingorder) Then
                            '            ' nvcFieldList1 = "UPDATE M04Cutoff set M03Orderqty=" & strOrderqty & ",M03Yarnstock='" & strYarncode & "',M03YarnType='" & strYarntype & "',M03IType='" & strI_Type & "',M03LineItem='" & strLineItem & "',M0330Class='" & str30class & "',M03Root='" & strRoot & "',M03MCNo='" & strMC & "',M03CuttingLine='" & strCutting_line & "',M03NoofRoll='" & strRolls & "' where M03OrderNo='" & strOrder & "' and M03Quality='" & strQuality & "' and M03Material='" & strMaterial & "'"
                            '            'ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                            '        Else
                            '            ncQryType = "ADD"
                            '            nvcFieldList1 = "(M02RefNo," & "M02OrderNo," & "M02Roll," & "M02SeqNo," & "M02Code," & "M02Spec," & "M02Dep," & "M02Date," & "M02Status," & "M02Class," & "M02Reason) " & "values(" & _RefNo & ",'" & strOrder & "'," & strCATRoll & ",'" & strSqNo & "','" & strCode & "'," & strSpec & ",'" & strDep & "','" & strDate & "','I','" & str30Class & "','" & strCode & "')"
                            '            'nvcFieldList1 = "(M04Order," & "M04CATRoll," & "M04SAPRoll) " & "values('" & Trim(strOrder) & "','" & strCATRoll & "','" & strSAPRoll & "')"
                            '            up_GetSetM02Defect(ncQryType, nvcFieldList1, nvcVccode, connection, transaction)


                            '            '---------------------------------------------------UPDATE P01PARAMETER TABLE
                            '            nvcFieldList1 = "UPDATE P01Parameter SET P01NO=P01NO +" & 1 & " WHERE P01CODE='DF'"
                            '            ExecuteNonQueryText(connection, transaction, nvcFieldList1)



                            '        End If
                            '    End If
                            'End If
                            linesList.RemoveAt(0)
                            ''  MsgBox(linesList.ToArray().ToString)
                            'IO.File.WriteAllLines(strFileName, linesList.ToArray())

                            strMtr = 0
                            strKg = 0

                    Else
                            Err.Raise(vbObjectError + 18001, "GenerateInstantFile(str,str,str)", "Invalid Record At Line " & CStr(lLineNo))
                    End If

                End If

                lLineNo = lLineNo + 1

            Loop

            transaction.Commit()
            DBEngin.CloseConnection(connection)
            ' transaction.Rollback()


            FileClose()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                MsgBox(lLineNo & "-" & strFileName)
            End If
        End Try
    End Function

    Function Upload_TobeDeliverd_File()
        Dim strFileName As String
        'strFileName = "\\Tjlapp04\grginspec_dload$\TJL CAT.txt"
        'strFileName = "E:\TJL_MILAN\SAP_DOWNLOADS\Sales Forcust\TobeDelivered.txt"
        ''Dim strFileName As String
        strFileName = ConfigurationManager.AppSettings("FilePath") + "\TobeDelivered.txt"
        Dim CurrGameWinningSerials(0 To MAX_SERIALS) As Long
        Dim fileHndl As Long
        Dim lLineNo As Long

        Dim strOrder, strLineItem, _
      strMerchant, strDis As String
        Dim strDep As String
        Dim strSpec As Double

        Dim strKg As Double
        Dim strMtr As Double
        ' Dim strFileName As String '= _
        Dim strDate As String
        ' Dim strDate As String
        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True

        Dim M03Knittingorder As DataSet
        Dim ncQryType As String
        Dim nvcVccode As String
        Dim linesList As New List(Of String)(IO.File.ReadAllLines(strFileName))
        Dim _RefNo As Integer
        Dim P01Parameter As DataSet
        Dim _Value As Double
        Dim M06Cls As DataSet
        Dim str30Class As String
        Dim strMaterial As String

        Dim nvcFieldList As String
        Dim nvcWhereClause As String
        Dim nvcVcLine As String
        Dim nvcVcDate As String
        Dim nvcQtype As String

        Dim characterToRemove As String
        Dim strQuality As String


        Try
            nvcFieldList1 = "delete from  M07TobeDelivered "
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            fileHndl = FreeFile()


            ' strFileName = Dir(strFileName)

            'UPGRADE_WARNING: Couldn't resolve default property of object fileHndl. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object strValidSerialFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            FileOpen(fileHndl, strFileName, OpenMode.Input)
            lLineNo = 0
            Dim strRow As String

            Do Until EOF(fileHndl)

                '  Line Input #fileHndl, strRow
                'UPGRADE_WARNING: Couldn't resolve default property of object fileHndl. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                strRow = LineInput(fileHndl)

                'If lLineNo = "1015" Then
                '    ' MsgBox("")
                'End If
                If Trim(strRow) <> "" Then
                    ncQryType = "LST"
                    If InStr(1, strRow, vbTab) > 0 Then

                        '  CurrGameWinningSerials(lLineNo) = Trim(Split(strRow, vbTab)(0))
                        strOrder = (Trim(Split(strRow, vbTab)(0)))
                        strLineItem = (Trim(Split(strRow, vbTab)(1)))
                        strDate = (Trim(Split(strRow, vbTab)(2)))
                        strDate = Microsoft.VisualBasic.Right(Microsoft.VisualBasic.Left(strDate, 6), 2)
                        strDate = strDate & "/" & Microsoft.VisualBasic.Right(Trim(Split(strRow, vbTab)(2)), 2)
                        strDate = strDate & "/" & Microsoft.VisualBasic.Left(Trim(Split(strRow, vbTab)(2)), 4)
                        strMaterial = (Trim(Split(strRow, vbTab)(3)))
                        strDis = (Trim(Split(strRow, vbTab)(4)))
                        strMerchant = (Trim(Split(strRow, vbTab)(10)))
                        str30Class = (Trim(Split(strRow, vbTab)(5)))
                        characterToRemove = "'"

                        strDis = (Replace(strDis, characterToRemove, ""))
                        str30Class = (Replace(str30Class, characterToRemove, ""))

                        characterToRemove = "-"

                        strMaterial = (Replace(strMaterial, characterToRemove, ""))

                        If IsNumeric(Microsoft.VisualBasic.Left(strDis, 1)) Then
                            strQuality = Microsoft.VisualBasic.Left(strDis, 5)
                        Else
                            strQuality = Microsoft.VisualBasic.Left(strDis, 8)
                        End If
                        M06Cls = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetM07TobeDelivered", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcCode", strOrder), New SqlParameter("@vcLine", strLineItem), New SqlParameter("@vcDate", strDate))
                        If isValidDataset(M06Cls) Then
                            For Each DTRow As DataRow In M06Cls.Tables(0).Rows
                                nvcWhereClause = "M07Sales_Order='" & strOrder & "' AND M07Line_Item='" & strLineItem & "' AND M07Date='" & strDate & "'"
                                ncQryType = "UPD"
                                nvcFieldList1 = "M07Qty_Kg='" & (Trim(Split(strRow, vbTab)(6))) & "',M07Qty_Mtr='" & (Trim(Split(strRow, vbTab)(7))) & "'"
                                up_GetSetM07TobeDelivered(ncQryType, nvcFieldList1, nvcWhereClause, nvcVcLine, nvcVcDate, nvcVccode, connection, transaction)
                                ' ExecuteNonQueryText(connection, transaction, "up_GetSetM06Delivary_Qty", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcCode", strOrder), New SqlParameter("@vcLine", strLineItem), New SqlParameter("@vcDate", strDate))
                            Next
                        Else
                            ncQryType = "ADD"
                            nvcFieldList1 = "(M07Sales_Order," & "M07Line_Item," & "M07Date," & "M07Material," & "M07Met_Dis," & "M07Retailer," & "M07Qty_Kg," & "M07Qty_Mtr," & "M07Unit_PriceKG," & "M07Unit_PriceMTR," & "M07Merchant," & "M07Quality) " & "values('" & Trim(strOrder) & "','" & strLineItem & "','" & strDate & "','" & strMaterial & "','" & strDis & "','" & str30Class & "','" & (Trim(Split(strRow, vbTab)(6))) & "','" & (Trim(Split(strRow, vbTab)(7))) & "','" & (Trim(Split(strRow, vbTab)(8))) & "','" & (Trim(Split(strRow, vbTab)(9))) & "','" & strMerchant & "','" & strQuality & "')"
                            up_GetSetM07TobeDelivered(ncQryType, nvcFieldList1, nvcWhereClause, nvcVcLine, nvcVcDate, nvcVccode, connection, transaction)
                        End If
                        '---------------------------------------------------------------------------------------------
                        '        nvcFieldList1 = "select * from M02Defect where M02OrderNo='" & strOrder & "' and M02Roll='" & strCATRoll & "' and M02SeqNo='" & strSqNo & "'"
                        '        M03Knittingorder = DbEngine.ExecuteDataset(connection, transaction, nvcFieldList1)
                        '        If isValidDataset(M03Knittingorder) Then
                        '            ' nvcFieldList1 = "UPDATE M04Cutoff set M03Orderqty=" & strOrderqty & ",M03Yarnstock='" & strYarncode & "',M03YarnType='" & strYarntype & "',M03IType='" & strI_Type & "',M03LineItem='" & strLineItem & "',M0330Class='" & str30class & "',M03Root='" & strRoot & "',M03MCNo='" & strMC & "',M03CuttingLine='" & strCutting_line & "',M03NoofRoll='" & strRolls & "' where M03OrderNo='" & strOrder & "' and M03Quality='" & strQuality & "' and M03Material='" & strMaterial & "'"
                        '            'ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                        '        Else
                        '            ncQryType = "ADD"
                        '            nvcFieldList1 = "(M02RefNo," & "M02OrderNo," & "M02Roll," & "M02SeqNo," & "M02Code," & "M02Spec," & "M02Dep," & "M02Date," & "M02Status," & "M02Class," & "M02Reason) " & "values(" & _RefNo & ",'" & strOrder & "'," & strCATRoll & ",'" & strSqNo & "','" & strCode & "'," & strSpec & ",'" & strDep & "','" & strDate & "','I','" & str30Class & "','" & strCode & "')"
                        '            'nvcFieldList1 = "(M04Order," & "M04CATRoll," & "M04SAPRoll) " & "values('" & Trim(strOrder) & "','" & strCATRoll & "','" & strSAPRoll & "')"
                        '            up_GetSetM02Defect(ncQryType, nvcFieldList1, nvcVccode, connection, transaction)


                        '            '---------------------------------------------------UPDATE P01PARAMETER TABLE
                        '            nvcFieldList1 = "UPDATE P01Parameter SET P01NO=P01NO +" & 1 & " WHERE P01CODE='DF'"
                        '            ExecuteNonQueryText(connection, transaction, nvcFieldList1)



                        '        End If
                        '    End If
                        'End If
                        linesList.RemoveAt(0)
                        ''  MsgBox(linesList.ToArray().ToString)
                        'IO.File.WriteAllLines(strFileName, linesList.ToArray())

                        strMtr = 0
                        strKg = 0

                    Else
                        Err.Raise(vbObjectError + 18001, "GenerateInstantFile(str,str,str)", "Invalid Record At Line " & CStr(lLineNo))
                    End If

                End If

                lLineNo = lLineNo + 1

            Loop

            transaction.Commit()
            DBEngin.CloseConnection(connection)
            ' transaction.Rollback()

            connection.Close()
            FileClose()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                MsgBox(lLineNo & "-" & strFileName)
                DBEngin.CloseConnection(connection)
                connection.Close()
            End If
        End Try
    End Function


    Function Upload_Delivary_File()
        Dim strFileName As String
        '  strFileName = "\\Tjlapp04\grginspec_dload$\TJL CAT.txt"
        ' strFileName = "E:\TJL_MILAN\SAP_DOWNLOADS\Sales Forcust\Delivered.txt"
        'Dim strFileName As String
        strFileName = ConfigurationManager.AppSettings("FilePath") + "\Delivered.txt"
        Dim CurrGameWinningSerials(0 To MAX_SERIALS) As Long
        Dim fileHndl As Long
        Dim lLineNo As Long

        Dim strOrder, strLineItem, _
      strMerchant, strDis As String
        Dim strDep As String
        Dim strSpec As Double

        Dim strKg As Double
        Dim strMtr As Double
        ' Dim strFileName As String '= _
        Dim strDate As String
        ' Dim strDate As String
        Dim nvcFieldList1 As String
        Dim strQuality As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True

        Dim M03Knittingorder As DataSet
        Dim ncQryType As String
        Dim nvcVccode As String
        Dim linesList As New List(Of String)(IO.File.ReadAllLines(strFileName))
        Dim _RefNo As Integer
        Dim P01Parameter As DataSet
        Dim _Value As Double
        Dim M06Cls As DataSet
        Dim str30Class As String
        Dim strMaterial As String

        Dim nvcFieldList As String
        Dim nvcWhereClause As String
        Dim nvcVcLine As String
        Dim nvcVcDate As String
        Dim nvcQtype As String

        Dim characterToRemove As String


        Try

            nvcFieldList1 = "delete from M06Delivary_Qty     "
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            fileHndl = FreeFile()


            ' strFileName = Dir(strFileName)

            'UPGRADE_WARNING: Couldn't resolve default property of object fileHndl. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object strValidSerialFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            FileOpen(fileHndl, strFileName, OpenMode.Input)
            lLineNo = 0
            Dim strRow As String

            Do Until EOF(fileHndl)

                '  Line Input #fileHndl, strRow
                'UPGRADE_WARNING: Couldn't resolve default property of object fileHndl. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                strRow = LineInput(fileHndl)


                If Trim(strRow) <> "" Then
                    ncQryType = "LST"
                    If InStr(1, strRow, vbTab) > 0 Then

                        '  CurrGameWinningSerials(lLineNo) = Trim(Split(strRow, vbTab)(0))
                        strOrder = (Trim(Split(strRow, vbTab)(0)))
                        strLineItem = (Trim(Split(strRow, vbTab)(1)))
                        strDate = (Trim(Split(strRow, vbTab)(2)))
                        strDate = Microsoft.VisualBasic.Right(Microsoft.VisualBasic.Left(strDate, 6), 2)
                        strDate = strDate & "/" & Microsoft.VisualBasic.Right(Trim(Split(strRow, vbTab)(2)), 2)
                        strDate = strDate & "/" & Microsoft.VisualBasic.Left(Trim(Split(strRow, vbTab)(2)), 4)
                        strMaterial = (Trim(Split(strRow, vbTab)(3)))
                        strDis = (Trim(Split(strRow, vbTab)(4)))
                        strMerchant = (Trim(Split(strRow, vbTab)(6)))
                        str30Class = (Trim(Split(strRow, vbTab)(5)))
                        characterToRemove = "'"
                        ' MsgBox(Trim(Split(strRow, vbTab)(10)))
                        strDis = (Replace(strDis, characterToRemove, ""))
                        str30Class = (Replace(str30Class, characterToRemove, ""))

                        characterToRemove = "-"
                        ' MsgBox(Trim(Split(strRow, vbTab)(10)))
                        strMaterial = (Replace(strMaterial, characterToRemove, ""))

                        If IsNumeric(Microsoft.VisualBasic.Left(strDis, 1)) Then
                            strQuality = Microsoft.VisualBasic.Left(strQuality, 5)
                        Else

                            strQuality = Microsoft.VisualBasic.Left(strQuality, 8)
                        End If
                        M06Cls = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetM06Delivary_Qty", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcCode", strOrder), New SqlParameter("@vcLine", strLineItem), New SqlParameter("@vcDate", strDate))
                        If isValidDataset(M06Cls) Then
                            For Each DTRow As DataRow In M06Cls.Tables(0).Rows
                                nvcWhereClause = "M06Sales_Order='" & strOrder & "' AND M06Line_Item='" & strLineItem & "' AND M06Date='" & strDate & "'"
                                ncQryType = "UPD"
                                nvcFieldList1 = "M06D_Qty_KG='" & (Trim(Split(strRow, vbTab)(7))) & "',M06D_Qty_Mtr='" & (Trim(Split(strRow, vbTab)(8))) & "'"
                                up_GetSetM06Delivary_Qty(ncQryType, nvcFieldList1, nvcWhereClause, nvcVcLine, nvcVcDate, nvcVccode, connection, transaction)
                                ' ExecuteNonQueryText(connection, transaction, "up_GetSetM06Delivary_Qty", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcCode", strOrder), New SqlParameter("@vcLine", strLineItem), New SqlParameter("@vcDate", strDate))
                            Next
                        Else
                            ncQryType = "ADD"
                            nvcFieldList1 = "(M06Sales_Order," & "M06Line_Item," & "M06Date," & "M06Material," & "M06Met_Dis," & "M06Retailer," & "M06Merchant," & "M06D_Qty_KG," & "M06D_Qty_Mtr," & "M06Unit_Kg," & "M06Unit_Mtr," & "M06Quality) " & "values('" & Trim(strOrder) & "','" & strLineItem & "','" & strDate & "','" & strMaterial & "','" & strDis & "','" & str30Class & "','" & strMerchant & "','" & (Trim(Split(strRow, vbTab)(7))) & "','" & (Trim(Split(strRow, vbTab)(8))) & "','" & (Trim(Split(strRow, vbTab)(10))) & "','" & (Trim(Split(strRow, vbTab)(9))) & "','" & strQuality & "')"
                            up_GetSetM06Delivary_Qty(ncQryType, nvcFieldList1, nvcWhereClause, nvcVcLine, nvcVcDate, nvcVccode, connection, transaction)
                        End If
                        '---------------------------------------------------------------------------------------------
                        '        nvcFieldList1 = "select * from M02Defect where M02OrderNo='" & strOrder & "' and M02Roll='" & strCATRoll & "' and M02SeqNo='" & strSqNo & "'"
                        '        M03Knittingorder = DbEngine.ExecuteDataset(connection, transaction, nvcFieldList1)
                        '        If isValidDataset(M03Knittingorder) Then
                        '            ' nvcFieldList1 = "UPDATE M04Cutoff set M03Orderqty=" & strOrderqty & ",M03Yarnstock='" & strYarncode & "',M03YarnType='" & strYarntype & "',M03IType='" & strI_Type & "',M03LineItem='" & strLineItem & "',M0330Class='" & str30class & "',M03Root='" & strRoot & "',M03MCNo='" & strMC & "',M03CuttingLine='" & strCutting_line & "',M03NoofRoll='" & strRolls & "' where M03OrderNo='" & strOrder & "' and M03Quality='" & strQuality & "' and M03Material='" & strMaterial & "'"
                        '            'ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                        '        Else
                        '            ncQryType = "ADD"
                        '            nvcFieldList1 = "(M02RefNo," & "M02OrderNo," & "M02Roll," & "M02SeqNo," & "M02Code," & "M02Spec," & "M02Dep," & "M02Date," & "M02Status," & "M02Class," & "M02Reason) " & "values(" & _RefNo & ",'" & strOrder & "'," & strCATRoll & ",'" & strSqNo & "','" & strCode & "'," & strSpec & ",'" & strDep & "','" & strDate & "','I','" & str30Class & "','" & strCode & "')"
                        '            'nvcFieldList1 = "(M04Order," & "M04CATRoll," & "M04SAPRoll) " & "values('" & Trim(strOrder) & "','" & strCATRoll & "','" & strSAPRoll & "')"
                        '            up_GetSetM02Defect(ncQryType, nvcFieldList1, nvcVccode, connection, transaction)


                        '            '---------------------------------------------------UPDATE P01PARAMETER TABLE
                        '            nvcFieldList1 = "UPDATE P01Parameter SET P01NO=P01NO +" & 1 & " WHERE P01CODE='DF'"
                        '            ExecuteNonQueryText(connection, transaction, nvcFieldList1)



                        '        End If
                        '    End If
                        'End If
                        linesList.RemoveAt(0)
                        ''  MsgBox(linesList.ToArray().ToString)
                        'IO.File.WriteAllLines(strFileName, linesList.ToArray())

                        strMtr = 0
                        strKg = 0

                    Else
                        Err.Raise(vbObjectError + 18001, "GenerateInstantFile(str,str,str)", "Invalid Record At Line " & CStr(lLineNo))
                    End If

                End If

                lLineNo = lLineNo + 1

            Loop

            transaction.Commit()
            DBEngin.CloseConnection(connection)
            ' transaction.Rollback()


            FileClose()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                MsgBox(lLineNo & "-" & strFileName)
            End If
        End Try
    End Function

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        Call Load_Gride_SalesOrder()
        chkReport.Checked = False
        cboMaterial.Text = ""
        cboMerch.Text = ""
        cboRetailer.Text = ""
        cboSales_Order.Text = ""
        chkUpload.Checked = False
    End Sub

    Private Sub chkReport_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkReport.CheckedChanged
        If chkReport.Checked = True Then
            Dim strMonth As String
            strMonth = MonthName(Month(txtFromDate.Text))
            Call Create_ReportDFCT(txtFromDate.Text, txtTodate.Text)
        End If
    End Sub

    Private Sub OPR0_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OPR0.Click

    End Sub
End Class