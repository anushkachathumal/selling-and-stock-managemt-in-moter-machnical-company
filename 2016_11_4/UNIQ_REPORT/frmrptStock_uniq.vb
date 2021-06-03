Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports CrystalDecisions.CrystalReports.Engine
Public Class frmrptStock_uniq
    Dim _Print_Status As String
    Dim c_dataCustomer1 As DataTable
    Dim c_dataCustomer2 As DataTable
    Dim _Dis As String
    Dim _From As Date
    Dim _To As Date
    Dim _Total_Cost As Double
    Dim _Total_Rate As Double
    Private Sub ExitToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem1.Click
        Me.Close()
    End Sub

    Function Load_Grid_Category()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select  ROW_NUMBER() OVER (ORDER BY s01item_code ) as  ##,max(M01Description)as [Category Name],max(M05Item_Code) as [Item Code],max(M05Brand_Name) as [Brand Name],MAX(tmpDescription) as [Item Name],SUM(Qty) as [Qty] from View_Stock_Balance inner join View_Product_Item on M05Ref_No=s01item_code  where m05status='A' and M01Description='" & Trim(cboCategory.Text) & "' group by s01item_code order by s01item_code "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid2.DataSource = M01
            UltraGrid2.Rows.Band.Columns(0).Width = 30
            UltraGrid2.Rows.Band.Columns(1).Width = 110
            UltraGrid2.Rows.Band.Columns(2).Width = 90
            UltraGrid2.Rows.Band.Columns(3).Width = 110
            UltraGrid2.Rows.Band.Columns(4).Width = 260
            UltraGrid2.Rows.Band.Columns(5).Width = 80
            'UltraGrid2.Rows.Band.Columns(6).Width = 110
            UltraGrid2.Rows.Band.Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            ' UltraGrid2.Rows.Band.Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            ' UltraGrid1.Rows.Band.Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

            con.ClearAllPools()
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
                con.close()

            End If
        End Try
    End Function

    Function Load_Grid_BrandName()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select  ROW_NUMBER() OVER (ORDER BY s01item_code ) as  ##,max(M01Description)as [Category Name],MAX(M05Item_Code) as [Item Code],max(M05Brand_Name) as [Brand Name],MAX(tmpDescription) as [Item Name],SUM(Qty) as [Qty] from View_Stock_Balance inner join View_Product_Item on M05Ref_No=s01item_code  where m05status='A' and M05Brand_Name='" & Trim(cboCategory.Text) & "' group by s01item_code order by s01item_code "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid2.DataSource = M01
            UltraGrid2.Rows.Band.Columns(0).Width = 30
            UltraGrid2.Rows.Band.Columns(1).Width = 110
            UltraGrid2.Rows.Band.Columns(2).Width = 90
            UltraGrid2.Rows.Band.Columns(3).Width = 110
            UltraGrid2.Rows.Band.Columns(4).Width = 260
            UltraGrid2.Rows.Band.Columns(5).Width = 80
            'UltraGrid2.Rows.Band.Columns(6).Width = 110
            UltraGrid2.Rows.Band.Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            ' UltraGrid2.Rows.Band.Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            ' UltraGrid1.Rows.Band.Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

            con.ClearAllPools()
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
                con.close()

            End If
        End Try
    End Function

    Function Load_Grid_Stock()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select  ROW_NUMBER() OVER (ORDER BY s01item_code ) as  ##,max(M01Description)as [Category Name],MAX(M05Item_Code) as [Item Code],max(M05Brand_Name) as [Brand Name],MAX(tmpDescription) as [Item Name],SUM(Qty) as [Qty] from View_Stock_Balance inner join View_Product_Item on M05Ref_No=s01item_code where m05status='A' group by s01item_code order by s01item_code  "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid2.DataSource = M01
            UltraGrid2.Rows.Band.Columns(0).Width = 30
            UltraGrid2.Rows.Band.Columns(1).Width = 110
            UltraGrid2.Rows.Band.Columns(2).Width = 90
            UltraGrid2.Rows.Band.Columns(3).Width = 110
            UltraGrid2.Rows.Band.Columns(4).Width = 260
            UltraGrid2.Rows.Band.Columns(5).Width = 80
            'UltraGrid2.Rows.Band.Columns(6).Width = 110
            UltraGrid2.Rows.Band.Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            ' UltraGrid2.Rows.Band.Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            ' UltraGrid1.Rows.Band.Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

            con.ClearAllPools()
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
                con.close()

            End If
        End Try
    End Function

    Function Load_Grid_nEGATIVE_Stock()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select  ROW_NUMBER() OVER (ORDER BY s01item_code ) as  ##,max(M01Description)as [Category Name],max(M05Item_Code) as [Item Code],max(M05Brand_Name) as [Brand Name],MAX(tmpDescription) as [Item Name],SUM(Qty) as [Qty] from View_Stock_Balance inner join View_Product_Item on M05Ref_No=s01item_code  where m05status='A'  group by s01item_code having SUM(Qty)<0 order by s01item_code "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid2.DataSource = M01
            UltraGrid2.Rows.Band.Columns(0).Width = 30
            UltraGrid2.Rows.Band.Columns(1).Width = 110
            UltraGrid2.Rows.Band.Columns(2).Width = 90
            UltraGrid2.Rows.Band.Columns(3).Width = 110
            UltraGrid2.Rows.Band.Columns(4).Width = 260
            UltraGrid2.Rows.Band.Columns(5).Width = 80
            'UltraGrid2.Rows.Band.Columns(6).Width = 110
            UltraGrid2.Rows.Band.Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            ' UltraGrid2.Rows.Band.Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            ' UltraGrid1.Rows.Band.Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

            con.ClearAllPools()
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
                con.close()

            End If
        End Try
    End Function


    Function Load_Gride_Stock_Movement1()
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
        Dim _Date As Date
        Dim X As Integer
        Dim _coloumCount As Integer
        Dim Value As Double
        Dim _STSting As String
        Dim _week As Integer
        Dim _BALANCE As Double

        Try

            UltraGrid2.DisplayLayout.Bands(0).Groups.Clear()
            UltraGrid2.DisplayLayout.Bands(0).Columns.Dispose()

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
            agroup1 = UltraGrid2.DisplayLayout.Bands(0).Groups.Add("")
         

            agroup1.Width = 110
            Dim dt As DataTable = New DataTable()
            ' dt.Columns.Add("ID", GetType(Integer))
            Dim colWork As New DataColumn("##", GetType(String))
            dt.Columns.Add(colWork)
            colWork.ReadOnly = True

            colWork = New DataColumn("Part No", GetType(String))
            colWork.MaxLength = 250
            dt.Columns.Add(colWork)
            colWork.ReadOnly = True

            colWork = New DataColumn("Item Name", GetType(String))
            colWork.MaxLength = 250
            dt.Columns.Add(colWork)
            colWork.ReadOnly = True

            'dt.Columns.Add("##", GetType(String))
            ' dt.Columns.Add("Shade", GetType(String))
            I = 0
            vcWhere = "select * from View_Product_Item where M05Status='A'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, vcWhere)
            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                dt.Rows.Add(M01.Tables(0).Rows(I)("M05Ref_No"), M01.Tables(0).Rows(I)("M05Item_Code"), UCase(M01.Tables(0).Rows(I)("M05Description")))
                I = I + 1
            Next

            Me.UltraGrid2.SetDataBinding(dt, Nothing)
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns(0).Group = agroup1
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns(1).Group = agroup1
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns(2).Group = agroup1
            ' Me.dg_YDP_Projection.DisplayLayout.Bands(0).Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns(2).Width = 260
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns(1).Width = 70
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns(0).Width = 50
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            Dim _Group As String
            'agroup2.Key = ""
            'agroup3.Key = ""
            'agroup4.Key = ""
            '' agroup5.Key = ""

            'I = 0
            ''T01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            'vcWhere = "T15Sales_Order='" & strSales_Order & "' and T15Line_Item=" & txtLine_Item.Text & ""
            'M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TPR"), New SqlParameter("@vcWhereClause1", vcWhere))
            '_week = M01.Tables(0).Rows.Count
            'For Each DTRow3 As DataRow In M01.Tables(0).Rows
            '    _Group = "Group" & I + 1
            '    If I = 0 Then
            '        'agroup2.Key = ""
            agroup2 = UltraGrid2.DisplayLayout.Bands(0).Groups.Add("Group1")

            agroup2.Header.Caption = "Stock Movement"
            agroup2.Width = 220

            Me.UltraGrid2.DisplayLayout.Bands(0).Columns.Add("O/B", "O/B on " & Year(txtDate3.Text) & "/" & Month(txtDate3.Text) & "/" & Microsoft.VisualBasic.Day(txtDate3.Text))
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns("O/B").Group = agroup2
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns("O/B").Width = 110
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns("O/B").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

            Me.UltraGrid2.DisplayLayout.Bands(0).Columns.Add("Item Issue", "Item Issue")
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns("Item Issue").Group = agroup2
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns("Item Issue").Width = 70
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns("Item Issue").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

            Me.UltraGrid2.DisplayLayout.Bands(0).Columns.Add("Direct Sales", "Direct Sales")
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns("Direct Sales").Group = agroup2
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns("Direct Sales").Width = 70
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns("Direct Sales").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

            Me.UltraGrid2.DisplayLayout.Bands(0).Columns.Add("GRN", "GRN")
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns("GRN").Group = agroup2
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns("GRN").Width = 70
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns("GRN").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center


            Me.UltraGrid2.DisplayLayout.Bands(0).Columns.Add("Sup_Return", "Sup_Return")
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns("Sup_Return").Group = agroup2
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns("Sup_Return").Width = 70
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns("Sup_Return").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

            Me.UltraGrid2.DisplayLayout.Bands(0).Columns.Add("Wastage", "Wastage")
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns("Wastage").Group = agroup2
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns("Wastage").Width = 70
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns("Wastage").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

            Me.UltraGrid2.DisplayLayout.Bands(0).Columns.Add("Balance", "Balance")
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns("Balance").Group = agroup2
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns("Balance").Width = 90
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns("Balance").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            I = 0

            For Each uRow As UltraGridRow In UltraGrid2.Rows
                _BALANCE = 0
                'OPANING BALANACE
                Sql = "SELECT SUM(S01Qty) AS S01Qty FROM S01Stock_Balance WHERE S01Item_Code='" & Trim(UltraGrid2.Rows(I).Cells(0).Text) & "' AND S01Date<'" & txtDate3.Text & "' AND S01Status='A' GROUP BY S01Item_Code"
                M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(M01) Then
                    UltraGrid2.Rows(I).Cells(3).Value = M01.Tables(0).Rows(0)("S01Qty")
                    _BALANCE = M01.Tables(0).Rows(0)("S01Qty")
                End If
                '===========================================================================
                'ITEM ISSUE
                Sql = "SELECT SUM(S01Qty) AS S01Qty FROM S01Stock_Balance WHERE S01Item_Code='" & Trim(UltraGrid2.Rows(I).Cells(0).Text) & "' AND S01Date BETWEEN '" & txtDate3.Text & "' AND '" & txtDate4.Text & "' AND S01Status='A'  and S01Tr_Type='ISSUE_ITEM' GROUP BY S01Item_Code"
                M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(M01) Then
                    UltraGrid2.Rows(I).Cells(4).Value = M01.Tables(0).Rows(0)("S01Qty")
                    _BALANCE = _BALANCE + CDbl(M01.Tables(0).Rows(0)("S01Qty"))
                End If
                '==========================================================================
                'DIRECT SALES
                Sql = "SELECT SUM(S01Qty) AS S01Qty FROM S01Stock_Balance WHERE S01Item_Code='" & Trim(UltraGrid2.Rows(I).Cells(0).Text) & "' AND S01Date BETWEEN '" & txtDate3.Text & "' AND '" & txtDate4.Text & "' AND S01Status='A'  and S01Tr_Type='DIRECT_SALES' GROUP BY S01Item_Code"
                M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(M01) Then
                    UltraGrid2.Rows(I).Cells(5).Value = M01.Tables(0).Rows(0)("S01Qty")
                    _BALANCE = _BALANCE + CDbl(M01.Tables(0).Rows(0)("S01Qty"))
                End If
                '=========================================================================
                'GRN
                Sql = "SELECT SUM(S01Qty) AS S01Qty FROM S01Stock_Balance WHERE S01Item_Code='" & Trim(UltraGrid2.Rows(I).Cells(0).Text) & "' AND S01Date BETWEEN '" & txtDate3.Text & "' AND '" & txtDate4.Text & "' AND S01Status='A'  and S01Tr_Type='GRN' GROUP BY S01Item_Code"
                M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(M01) Then
                    UltraGrid2.Rows(I).Cells(6).Value = M01.Tables(0).Rows(0)("S01Qty")
                    _BALANCE = _BALANCE + CDbl(M01.Tables(0).Rows(0)("S01Qty"))
                End If
                '=========================================================================
                'SUPPLIER RETURN
                Sql = "SELECT SUM(S01Qty) AS S01Qty FROM S01Stock_Balance WHERE S01Item_Code='" & Trim(UltraGrid2.Rows(I).Cells(0).Text) & "' AND S01Date BETWEEN '" & txtDate3.Text & "' AND '" & txtDate4.Text & "' AND S01Status='A'  and S01Tr_Type='SP_RETURN' GROUP BY S01Item_Code"
                M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(M01) Then
                    UltraGrid2.Rows(I).Cells(7).Value = M01.Tables(0).Rows(0)("S01Qty")
                    _BALANCE = _BALANCE + CDbl(M01.Tables(0).Rows(0)("S01Qty"))
                End If
                '========================================================================
                'WASTAGE
                Sql = "SELECT SUM(S01Qty) AS S01Qty FROM S01Stock_Balance WHERE S01Item_Code='" & Trim(UltraGrid2.Rows(I).Cells(0).Text) & "' AND S01Date BETWEEN '" & txtDate3.Text & "' AND '" & txtDate4.Text & "' AND S01Status='A'  and S01Tr_Type='WST' GROUP BY S01Item_Code"
                M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(M01) Then
                    UltraGrid2.Rows(I).Cells(8).Value = M01.Tables(0).Rows(0)("S01Qty")
                    _BALANCE = _BALANCE + CDbl(M01.Tables(0).Rows(0)("S01Qty"))
                End If
                UltraGrid2.Rows(I).Cells(9).Value = _BALANCE
                I = I + 1
            Next


            con.ClearAllPools()
            con.close()

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
                con.close()

            End If
        End Try
    End Function

    Function Load_Gride_Stock_Movement2()
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
        Dim _Date As Date
        Dim X As Integer
        Dim _coloumCount As Integer
        Dim Value As Double
        Dim _STSting As String
        Dim _week As Integer
        Dim _BALANCE As Double

        Try

            UltraGrid2.DisplayLayout.Bands(0).Groups.Clear()
            UltraGrid2.DisplayLayout.Bands(0).Columns.Dispose()

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
            agroup1 = UltraGrid2.DisplayLayout.Bands(0).Groups.Add("")


            agroup1.Width = 110
            Dim dt As DataTable = New DataTable()
            ' dt.Columns.Add("ID", GetType(Integer))
            Dim colWork As New DataColumn("##", GetType(String))
            dt.Columns.Add(colWork)
            colWork.ReadOnly = True
            colWork = New DataColumn("Item Name", GetType(String))
            colWork.MaxLength = 250
            dt.Columns.Add(colWork)
            colWork.ReadOnly = True

            'dt.Columns.Add("##", GetType(String))
            ' dt.Columns.Add("Shade", GetType(String))
            I = 0
            vcWhere = "select * from View_Product_Item where M05Status='A' and M01Description='" & Trim(cboCategory1.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, vcWhere)
            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                dt.Rows.Add(M01.Tables(0).Rows(I)("M05Item_Code"), UCase(M01.Tables(0).Rows(I)("M05Description")))
                I = I + 1
            Next

            Me.UltraGrid2.SetDataBinding(dt, Nothing)
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns(0).Group = agroup1
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns(1).Group = agroup1
            ' Me.dg_YDP_Projection.DisplayLayout.Bands(0).Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns(1).Width = 260
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns(0).Width = 70
            Dim _Group As String
            'agroup2.Key = ""
            'agroup3.Key = ""
            'agroup4.Key = ""
            '' agroup5.Key = ""

            'I = 0
            ''T01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            'vcWhere = "T15Sales_Order='" & strSales_Order & "' and T15Line_Item=" & txtLine_Item.Text & ""
            'M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TPR"), New SqlParameter("@vcWhereClause1", vcWhere))
            '_week = M01.Tables(0).Rows.Count
            'For Each DTRow3 As DataRow In M01.Tables(0).Rows
            '    _Group = "Group" & I + 1
            '    If I = 0 Then
            '        'agroup2.Key = ""
            agroup2 = UltraGrid2.DisplayLayout.Bands(0).Groups.Add("Group1")

            agroup2.Header.Caption = "Stock Movement"
            agroup2.Width = 220

            Me.UltraGrid2.DisplayLayout.Bands(0).Columns.Add("O/B", "O/B on " & Year(txtD1.Text) & "/" & Month(txtD1.Text) & "/" & Microsoft.VisualBasic.Day(txtD1.Text))
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns("O/B").Group = agroup2
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns("O/B").Width = 110
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns("O/B").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

            Me.UltraGrid2.DisplayLayout.Bands(0).Columns.Add("Item Issue", "Item Issue")
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns("Item Issue").Group = agroup2
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns("Item Issue").Width = 70
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns("Item Issue").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

            Me.UltraGrid2.DisplayLayout.Bands(0).Columns.Add("Direct Sales", "Direct Sales")
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns("Direct Sales").Group = agroup2
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns("Direct Sales").Width = 70
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns("Direct Sales").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

            Me.UltraGrid2.DisplayLayout.Bands(0).Columns.Add("GRN", "GRN")
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns("GRN").Group = agroup2
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns("GRN").Width = 70
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns("GRN").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center


            Me.UltraGrid2.DisplayLayout.Bands(0).Columns.Add("Sup_Return", "Sup_Return")
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns("Sup_Return").Group = agroup2
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns("Sup_Return").Width = 70
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns("Sup_Return").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

            Me.UltraGrid2.DisplayLayout.Bands(0).Columns.Add("Wastage", "Wastage")
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns("Wastage").Group = agroup2
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns("Wastage").Width = 70
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns("Wastage").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

            Me.UltraGrid2.DisplayLayout.Bands(0).Columns.Add("Balance", "Balance")
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns("Balance").Group = agroup2
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns("Balance").Width = 90
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns("Balance").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            I = 0

            For Each uRow As UltraGridRow In UltraGrid2.Rows
                _BALANCE = 0
                'OPANING BALANACE
                Sql = "SELECT SUM(S01Qty) AS S01Qty FROM S01Stock_Balance WHERE S01Item_Code='" & Trim(UltraGrid2.Rows(I).Cells(0).Text) & "' AND S01Date<'" & txtD1.Text & "' AND S01Status='A' GROUP BY S01Item_Code"
                M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(M01) Then
                    UltraGrid2.Rows(I).Cells(2).Value = M01.Tables(0).Rows(0)("S01Qty")
                    _BALANCE = M01.Tables(0).Rows(0)("S01Qty")
                End If
                '===========================================================================
                'ITEM ISSUE
                Sql = "SELECT SUM(S01Qty) AS S01Qty FROM S01Stock_Balance WHERE S01Item_Code='" & Trim(UltraGrid2.Rows(I).Cells(0).Text) & "' AND S01Date BETWEEN '" & txtD1.Text & "' AND '" & txtD2.Text & "' AND S01Status='A'  and S01Tr_Type='ISSUE_ITEM' GROUP BY S01Item_Code"
                M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(M01) Then
                    UltraGrid2.Rows(I).Cells(3).Value = M01.Tables(0).Rows(0)("S01Qty")
                    _BALANCE = _BALANCE + CDbl(M01.Tables(0).Rows(0)("S01Qty"))
                End If
                '==========================================================================
                'DIRECT SALES
                Sql = "SELECT SUM(S01Qty) AS S01Qty FROM S01Stock_Balance WHERE S01Item_Code='" & Trim(UltraGrid2.Rows(I).Cells(0).Text) & "' AND S01Date BETWEEN '" & txtD1.Text & "' AND '" & txtD2.Text & "' AND S01Status='A'  and S01Tr_Type='DIRECT_SALES' GROUP BY S01Item_Code"
                M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(M01) Then
                    UltraGrid2.Rows(I).Cells(4).Value = M01.Tables(0).Rows(0)("S01Qty")
                    _BALANCE = _BALANCE + CDbl(M01.Tables(0).Rows(0)("S01Qty"))
                End If
                '=========================================================================
                'GRN
                Sql = "SELECT SUM(S01Qty) AS S01Qty FROM S01Stock_Balance WHERE S01Item_Code='" & Trim(UltraGrid2.Rows(I).Cells(0).Text) & "' AND S01Date BETWEEN '" & txtD1.Text & "' AND '" & txtD2.Text & "' AND S01Status='A'  and S01Tr_Type='GRN' GROUP BY S01Item_Code"
                M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(M01) Then
                    UltraGrid2.Rows(I).Cells(5).Value = M01.Tables(0).Rows(0)("S01Qty")
                    _BALANCE = _BALANCE + CDbl(M01.Tables(0).Rows(0)("S01Qty"))
                End If
                '=========================================================================
                'SUPPLIER RETURN
                Sql = "SELECT SUM(S01Qty) AS S01Qty FROM S01Stock_Balance WHERE S01Item_Code='" & Trim(UltraGrid2.Rows(I).Cells(0).Text) & "' AND S01Date BETWEEN '" & txtD1.Text & "' AND '" & txtD2.Text & "' AND S01Status='A'  and S01Tr_Type='SP_RETURN' GROUP BY S01Item_Code"
                M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(M01) Then
                    UltraGrid2.Rows(I).Cells(6).Value = M01.Tables(0).Rows(0)("S01Qty")
                    _BALANCE = _BALANCE + CDbl(M01.Tables(0).Rows(0)("S01Qty"))
                End If
                '========================================================================
                'WASTAGE
                Sql = "SELECT SUM(S01Qty) AS S01Qty FROM S01Stock_Balance WHERE S01Item_Code='" & Trim(UltraGrid2.Rows(I).Cells(0).Text) & "' AND S01Date BETWEEN '" & txtD1.Text & "' AND '" & txtD2.Text & "' AND S01Status='A'  and S01Tr_Type='WST' GROUP BY S01Item_Code"
                M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(M01) Then
                    UltraGrid2.Rows(I).Cells(7).Value = M01.Tables(0).Rows(0)("S01Qty")
                    _BALANCE = _BALANCE + CDbl(M01.Tables(0).Rows(0)("S01Qty"))
                End If
                UltraGrid2.Rows(I).Cells(8).Value = _BALANCE
                I = I + 1
            Next


            con.ClearAllPools()
            con.close()

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
                con.close()

            End If
        End Try
    End Function

    Function Load_Gride()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTable_StockItemsrpt
        UltraGrid2.DataSource = c_dataCustomer1
        With UltraGrid2
            .DisplayLayout.Bands(0).Columns(0).Width = 130
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 180
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(3).Width = 80
            .DisplayLayout.Bands(0).Columns(3).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(4).Width = 90
            .DisplayLayout.Bands(0).Columns(4).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(5).Width = 90
            .DisplayLayout.Bands(0).Columns(5).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(6).Width = 110
            .DisplayLayout.Bands(0).Columns(6).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(7).Width = 110
            .DisplayLayout.Bands(0).Columns(7).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            
        End With
    End Function

    Function Load_Gride_Brand()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTable_StockItemsrpt_Brand
        UltraGrid2.DataSource = c_dataCustomer1
        With UltraGrid2
            .DisplayLayout.Bands(0).Columns(0).Width = 30
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 130
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 110
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(3).Width = 80
            .DisplayLayout.Bands(0).Columns(3).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(4).Width = 180
            .DisplayLayout.Bands(0).Columns(4).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(5).Width = 60
            .DisplayLayout.Bands(0).Columns(5).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(6).Width = 80
            .DisplayLayout.Bands(0).Columns(6).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(7).Width = 80
            .DisplayLayout.Bands(0).Columns(7).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(8).Width = 100
            .DisplayLayout.Bands(0).Columns(8).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(8).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(9).Width = 100
            .DisplayLayout.Bands(0).Columns(9).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(9).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
        End With
    End Function


    Function Load_Grid_Stock_Value()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Dim I As Integer
        Dim _St As String

        Dim Value As Double
        Dim _Rowcount As Integer

        Try
            Sql = "select  ROW_NUMBER() OVER (ORDER BY s01item_code ) as  ##,max(M01Description)as Category_Name,MAX(M05Item_Code) AS M05Item_Code ,max(M05Brand_Name) as Brand_Name,MAX(tmpDescription) as Item_Name,SUM(Qty) as Qty,max(M05Cost) as Cost,max(M05Retail)  as Rate,max(M05Cost)*SUM(Qty) as Total_Cost,max(M05Retail)*SUM(Qty)  as Total_Rate from View_Stock_Balance inner join View_Product_Item on M05Ref_No=s01item_code where m05status='A' group by s01item_code order by s01item_code "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            I = 0
            _Total_Cost = 0
            _Total_Rate = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow
                'newRow("##") = M01.Tables(0).Rows(I)("##")
                newRow("Category") = Trim(M01.Tables(0).Rows(I)("Category_Name"))
                '  newRow("Brand Name") = Trim(M01.Tables(0).Rows(I)("Brand_Name"))
                newRow("Item Code") = Trim(M01.Tables(0).Rows(I)("M05Item_Code"))
                newRow("Item Name") = Trim(M01.Tables(0).Rows(I)("Item_Name"))
                'Value = M01.Tables(0).Rows(I)("Qty")
                '_St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                '_St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Qty") = CInt(M01.Tables(0).Rows(I)("Qty"))
                Value = M01.Tables(0).Rows(I)("Cost")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("#Unit Cost") = _St
                Value = M01.Tables(0).Rows(I)("Rate")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("#Unit Rate") = _St
                Value = M01.Tables(0).Rows(I)("Total_Cost")
                _Total_Cost = _Total_Cost + Value
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("#Total Cost") = _St


                Value = M01.Tables(0).Rows(I)("Total_Rate")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("#Total Rate") = _St
                _Total_Rate = _Total_Rate + Value
                c_dataCustomer1.Rows.Add(newRow)

                I = I + 1
            Next

            Dim newRow1 As DataRow = c_dataCustomer1.NewRow
            newRow1("Category") = ""
            c_dataCustomer1.Rows.Add(newRow1)

            Dim newRow2 As DataRow = c_dataCustomer1.NewRow
            _St = (_Total_Cost.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", _Total_Cost))
            newRow2("#Total Cost") = _St

            _St = (_Total_Rate.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", _Total_Rate))
            newRow2("#Total Rate") = _St

            c_dataCustomer1.Rows.Add(newRow2)

            _Rowcount = UltraGrid2.Rows.Count - 1
            UltraGrid2.Rows(_Rowcount).Cells(6).Appearance.BackColor = Color.Gold
            UltraGrid2.Rows(_Rowcount).Cells(7).Appearance.BackColor = Color.Gold
            UltraGrid2.Rows(_Rowcount).Cells(6).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            UltraGrid2.Rows(_Rowcount).Cells(7).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            con.ClearAllPools()
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
                con.close()

            End If
        End Try
    End Function

    Function Load_Grid_Stock_Category()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Dim I As Integer
        Dim _St As String

        Dim Value As Double
        Dim _Rowcount As Integer

        Try
            Sql = "select  ROW_NUMBER() OVER (ORDER BY s01item_code ) as  ##,max(M01Description)as Category_Name,max(M05Item_Code) as M05Item_Code ,max(M05Brand_Name) as Brand_Name,MAX(tmpDescription) as Item_Name,SUM(Qty) as Qty,max(M05Cost) as Cost,max(M05Retail)  as Rate,max(M05Cost)*SUM(Qty) as Total_Cost,max(M05Retail)*SUM(Qty)  as Total_Rate from View_Stock_Balance inner join View_Product_Item on M05Ref_No=s01item_code  where m05status='A' and M01Description='" & Trim(cboCategory.Text) & "' group by s01item_code order by s01item_code "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            I = 0
            _Total_Cost = 0
            _Total_Rate = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow
                newRow("##") = M01.Tables(0).Rows(I)("##")
                newRow("Category") = Trim(M01.Tables(0).Rows(I)("Category_Name"))
                newRow("Brand Name") = Trim(M01.Tables(0).Rows(I)("Brand_Name"))
                newRow("Item Code") = Trim(M01.Tables(0).Rows(I)("M05Item_Code"))
                newRow("Item Name") = Trim(M01.Tables(0).Rows(I)("Item_Name"))
                'Value = M01.Tables(0).Rows(I)("Qty")
                '_St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                '_St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Qty") = CInt(M01.Tables(0).Rows(I)("Qty"))
                Value = M01.Tables(0).Rows(I)("Cost")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("#Unit Cost") = _St
                Value = M01.Tables(0).Rows(I)("Rate")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("#Unit Rate") = _St
                Value = M01.Tables(0).Rows(I)("Total_Cost")
                _Total_Cost = _Total_Cost + Value
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("#Total Cost") = _St


                Value = M01.Tables(0).Rows(I)("Total_Rate")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("#Total Rate") = _St
                _Total_Rate = _Total_Rate + Value
                c_dataCustomer1.Rows.Add(newRow)

                I = I + 1
            Next

            Dim newRow1 As DataRow = c_dataCustomer1.NewRow
            newRow1("Category") = ""
            c_dataCustomer1.Rows.Add(newRow1)

            Dim newRow2 As DataRow = c_dataCustomer1.NewRow
            _St = (_Total_Cost.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", _Total_Cost))
            newRow2("#Total Cost") = _St

            _St = (_Total_Rate.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", _Total_Rate))
            newRow2("#Total Rate") = _St

            c_dataCustomer1.Rows.Add(newRow2)

            _Rowcount = UltraGrid2.Rows.Count - 1
            UltraGrid2.Rows(_Rowcount).Cells(9).Appearance.BackColor = Color.Gold
            UltraGrid2.Rows(_Rowcount).Cells(8).Appearance.BackColor = Color.Gold
            UltraGrid2.Rows(_Rowcount).Cells(8).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            UltraGrid2.Rows(_Rowcount).Cells(9).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            con.ClearAllPools()
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
                con.close()

            End If
        End Try
    End Function

    Function Load_Grid_Stock_Brand_Name()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Dim I As Integer
        Dim _St As String

        Dim Value As Double
        Dim _Rowcount As Integer

        Try
            Sql = "select  ROW_NUMBER() OVER (ORDER BY s01item_code ) as  ##,max(M01Description)as Category_Name,max(M05Item_Code) as M05Item_Code ,max(M05Brand_Name) as Brand_Name,MAX(tmpDescription) as Item_Name,SUM(Qty) as Qty,max(M05Cost) as Cost,max(M05Retail)  as Rate,max(M05Cost)*SUM(Qty) as Total_Cost,max(M05Retail)*SUM(Qty)  as Total_Rate from View_Stock_Balance inner join View_Product_Item on M05Ref_No=s01item_code  where m05status='A' and M05Brand_Name='" & Trim(cboCategory.Text) & "' group by s01item_code order by s01item_code "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            I = 0
            _Total_Cost = 0
            _Total_Rate = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow
                newRow("##") = M01.Tables(0).Rows(I)("##")
                newRow("Category") = Trim(M01.Tables(0).Rows(I)("Category_Name"))
                newRow("Brand Name") = Trim(M01.Tables(0).Rows(I)("Brand_Name"))
                newRow("Item Code") = Trim(M01.Tables(0).Rows(I)("M05Item_Code"))
                newRow("Item Name") = Trim(M01.Tables(0).Rows(I)("Item_Name"))
                'Value = M01.Tables(0).Rows(I)("Qty")
                '_St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                '_St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Qty") = CInt(M01.Tables(0).Rows(I)("Qty"))
                Value = M01.Tables(0).Rows(I)("Cost")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("#Unit Cost") = _St
                Value = M01.Tables(0).Rows(I)("Rate")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("#Unit Rate") = _St
                Value = M01.Tables(0).Rows(I)("Total_Cost")
                _Total_Cost = _Total_Cost + Value
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("#Total Cost") = _St


                Value = M01.Tables(0).Rows(I)("Total_Rate")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("#Total Rate") = _St
                _Total_Rate = _Total_Rate + Value
                c_dataCustomer1.Rows.Add(newRow)

                I = I + 1
            Next

            Dim newRow1 As DataRow = c_dataCustomer1.NewRow
            newRow1("##") = ""
            c_dataCustomer1.Rows.Add(newRow1)

            Dim newRow2 As DataRow = c_dataCustomer1.NewRow
            _St = (_Total_Cost.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", _Total_Cost))
            newRow2("#Total Cost") = _St

            _St = (_Total_Rate.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", _Total_Rate))
            newRow2("#Total Rate") = _St

            c_dataCustomer1.Rows.Add(newRow2)

            _Rowcount = UltraGrid2.Rows.Count - 1
            UltraGrid2.Rows(_Rowcount).Cells(9).Appearance.BackColor = Color.Gold
            UltraGrid2.Rows(_Rowcount).Cells(8).Appearance.BackColor = Color.Gold
            UltraGrid2.Rows(_Rowcount).Cells(8).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            UltraGrid2.Rows(_Rowcount).Cells(9).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            con.ClearAllPools()
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
                con.close()

            End If
        End Try
    End Function


    Private Sub frmrptStock_uniq_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_Grid_Stock()
        _Print_Status = "A1"
    End Sub

    Private Sub WithOutValuationToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WithOutValuationToolStripMenuItem.Click
        _Print_Status = "A1"
        Panel1.Visible = False
        Panel2.Visible = False
        Panel3.Visible = False
        Call Load_Grid_Stock()
    End Sub

    Private Sub RefreshToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RefreshToolStripMenuItem.Click
        Call Load_Grid_Stock()
        _Print_Status = "A1"
        Panel1.Visible = False
        Panel1.Visible = False
        Panel2.Visible = False
    End Sub

    Private Sub WithValuationToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        _Print_Status = "A2"
        Call Load_Grid_Stock_Value()
    End Sub

    Private Sub WithValuationToolStripMenuItem_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WithValuationToolStripMenuItem.Click
        _Print_Status = "A2"
        Panel1.Visible = False
        Panel2.Visible = False
        Panel3.Visible = False
        Call Load_Gride()
        Call Load_Grid_Stock_Value()

    End Sub

    Private Sub ToolStripMenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem3.Click
        Dim B As New ReportDocument
        Dim A As String
        Try
            If _Print_Status = "A2" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\Stock1.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "sainfinity")
                B.SetParameterValue("Total_cost", _Total_Cost)
                B.SetParameterValue("Total_Rate", _Total_Rate)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{M05Item_Master.M05Status} ='A' "
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _Print_Status = "A1" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\Stock.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "sainfinity")
                'B.SetParameterValue("To", _To)
                'B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{M05Item_Master.M05Status} ='A' "
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _Print_Status = "B1" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\Stock.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "sainfinity")
                'B.SetParameterValue("To", _To)
                'B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{M05Item_Master.M05Status} ='A' and {M01Category.M01Description}='" & _Dis & "' "
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _Print_Status = "B2" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\Stock1.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "sainfinity")
                B.SetParameterValue("Total_cost", _Total_Cost)
                B.SetParameterValue("Total_Rate", _Total_Rate)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{M05Item_Master.M05Status} ='A' and {M01Category.M01Description}='" & _Dis & "' "
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _Print_Status = "C1" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\Stock.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "sainfinity")
                'B.SetParameterValue("To", _To)
                'B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{M05Item_Master.M05Status} ='A' and {M05Item_Master.M05Brand_Name}='" & _Dis & "'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _Print_Status = "C2" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\Stock1.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "sainfinity")
                B.SetParameterValue("Total_cost", _Total_Cost)
                B.SetParameterValue("Total_Rate", _Total_Rate)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{M05Item_Master.M05Status} ='A' and {M05Item_Master.M05Brand_Name}='" & _Dis & "' "
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _Print_Status = "D" Then
                Call Save_Stock_Movement()
                A = ConfigurationManager.AppSettings("ReportPath") + "\Srock_Movement.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "sainfinity")
                B.SetParameterValue("From", _From)
                B.SetParameterValue("To", _To)
                B.SetParameterValue("OB", "O/B on " & Year(_From) & "/" & Month(_From) & "/" & Microsoft.VisualBasic.Day(_From))
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_Product_Item.M05Status} ='A' "
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _Print_Status = "D1" Then
                Call Save_Stock_Movement()
                A = ConfigurationManager.AppSettings("ReportPath") + "\Srock_Movement.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "sainfinity")
                B.SetParameterValue("From", _From)
                B.SetParameterValue("To", _To)
                B.SetParameterValue("OB", "O/B on " & Year(_From) & "/" & Month(_From) & "/" & Microsoft.VisualBasic.Day(_From))
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_Product_Item.M01Description}='" & _Dis & "' AND {View_Product_Item.M05Status}='A'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _Print_Status = "X1" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\Stock.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "sainfinity")
                'B.SetParameterValue("To", _To)
                'B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{M05Item_Master.M05Status} ='A' and {View_Stock_Balance.Qty} < 0"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            End If
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)


            End If
        End Try
    End Sub

    Function Load_Category()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet

        Try
            Sql = "select M01Description as [##] from M01Category WHERE M01Status='A' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboCategory
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 275

            End With

            With cboCategory1
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 275

            End With

            con.ClearAllPools()
            con.close()


        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
                con.close()

            End If
        End Try
    End Function

    Function Load_BrandName()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet

        Try
            Sql = "select M05Brand_Name as [##] from M05Item_Master WHERE M05Status='A' group by M05Brand_Name "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboCategory
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 275

            End With

            con.ClearAllPools()
            con.close()


        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
                con.close()

            End If
        End Try
    End Function

    Private Sub UltraButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton1.Click
        If _Print_Status = "B1" Then
            _Dis = Trim(cboCategory.Text)
            Call Load_Grid_Category()
            Panel1.Visible = False
        ElseIf _Print_Status = "B2" Then
            _Dis = Trim(cboCategory.Text)
            Call Load_Gride_Brand()
            Call Load_Grid_Stock_Category()
            Panel1.Visible = False
        ElseIf _Print_Status = "C1" Then
            _Dis = Trim(cboCategory.Text)
            Call Load_Grid_BrandName()
            Panel1.Visible = False
        ElseIf _Print_Status = "C2" Then
            _Dis = Trim(cboCategory.Text)
            Call Load_Gride_Brand()
            Call Load_Grid_Stock_Brand_Name()
            Panel1.Visible = False
        End If
    End Sub

    Private Sub WithoutValuationToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WithoutValuationToolStripMenuItem1.Click
        _Print_Status = "B1"
        Label1.Text = "Category"
        Call Load_Category()
        Panel1.Visible = True
        Panel2.Visible = False
        Panel3.Visible = False
        cboCategory.Text = ""
        cboCategory.ToggleDropdown()

    End Sub

    Private Sub WithValuationToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WithValuationToolStripMenuItem1.Click
        _Print_Status = "B2"
        Label1.Text = "Category"
        Call Load_Category()
        Panel1.Visible = True
        Panel2.Visible = False
        Panel3.Visible = False
        cboCategory.Text = ""
        cboCategory.ToggleDropdown()

    End Sub

    Private Sub WithoutValuationToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WithoutValuationToolStripMenuItem2.Click
        _Print_Status = "C1"
        Label1.Text = "Brand Name"
        Call Load_BrandName()
        Panel1.Visible = True
        Panel2.Visible = False
        Panel3.Visible = False
        cboCategory.Text = ""
        cboCategory.ToggleDropdown()
    End Sub

    Private Sub WithValuationToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WithValuationToolStripMenuItem2.Click
        _Print_Status = "C2"
        Label1.Text = "Brand Name"
        Call Load_BrandName()
        Panel1.Visible = True
        Panel2.Visible = False
        Panel3.Visible = False
        cboCategory.Text = ""
        cboCategory.ToggleDropdown()
    End Sub

    Private Sub StockAnalysisToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles StockAnalysisToolStripMenuItem.Click
       
    End Sub

    Private Sub UltraButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton2.Click
        If _Print_Status = "D" Then
            Me.Cursor = Cursors.WaitCursor
            _From = txtDate3.Text
            _To = txtDate4.Text
            Call Load_Gride_Stock_Movement1()
            Panel2.Visible = False
            Me.Cursor = Cursors.Arrow
        End If
    End Sub

    Function Save_Stock_Movement()
        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Dim MB51 As DataSet
        Dim i As Integer

        i = 0
        nvcFieldList1 = "DELETE FROM R01Report_Stock_Movement"
        ExecuteNonQueryText(connection, transaction, nvcFieldList1)

        For Each uRow As UltraGridRow In UltraGrid2.Rows
            nvcFieldList1 = "Insert Into R01Report_Stock_Movement(R01Item_Code,R01OB,R01issue,R01Sales,R01GRN,R01Return,R01Wastage,R01Balance)" & _
                                        " values('" & Trim(UltraGrid2.Rows(i).Cells(0).Text) & "','" & Trim(UltraGrid2.Rows(i).Cells(2).Text) & "', '" & Trim(UltraGrid2.Rows(i).Cells(3).Text) & "','" & Trim(UltraGrid2.Rows(i).Cells(4).Text) & "','" & Trim(UltraGrid2.Rows(i).Cells(5).Text) & "','" & Trim(UltraGrid2.Rows(i).Cells(6).Text) & "','" & Trim(UltraGrid2.Rows(i).Cells(7).Text) & "','" & Trim(UltraGrid2.Rows(i).Cells(8).Text) & "')"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            i = i + 1
        Next
        transaction.Commit()
        connection.ClearAllPools()
        connection.Close()
    End Function

    Private Sub ToolStripMenuItem6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem6.Click
        _Print_Status = "X1"
        Call Load_Grid_nEGATIVE_Stock()

    End Sub

    Private Sub AllStockToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AllStockToolStripMenuItem.Click
        'Call Load_Gride_Stock_Movement()
        Panel2.Visible = True
        Panel1.Visible = False
        Panel3.Visible = False
        txtDate3.Text = Today
        txtDate4.Text = Today
        _Print_Status = "D"
    End Sub

    Private Sub ByCategoryToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ByCategoryToolStripMenuItem.Click
        Panel2.Visible = False
        Panel1.Visible = False
        Panel3.Visible = True
        txtD1.Text = Today
        txtD2.Text = Today
        Call Load_Category()
        _Print_Status = "D1"
        cboCategory1.Text = ""
        cboCategory1.ToggleDropdown()
    End Sub

    Private Sub UltraButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton3.Click
        If _Print_Status = "D1" Then
            _From = txtD1.Text
            _To = txtD2.Text
            _Dis = Trim(cboCategory1.Text)
            Call Load_Gride_Stock_Movement2()
            Panel3.Visible = False
        End If
    End Sub

    Private Sub ToolStripMenuItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem4.Click
        _Print_Status = "Z1"
        Panel1.Visible = False
        Panel2.Visible = False
        Panel3.Visible = False
        Panel4.Visible = True
        txtC1.Text = Today
        txtC2.Text = Today
        Call Load_Items()
        cboItem.ToggleDropdown()
    End Sub

    Function Load_Items()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select M05Item_Code as [##],tmpDescription as [Item Name] from View_Product_Item where M05Status='A'  order by M05ID"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboItem
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 275
                .Rows.Band.Columns(1).Width = 310

            End With
            con.ClearAllPools()
            con.CLOSE()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
                con.CLOSE()
            End If
        End Try
    End Function


    Private Sub UltraButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton4.Click
        _From = txtC1.Text
        _To = txtC2.Text
        _Dis = Trim(cboItem.Text)
        Call Load_Gride_Items_Movement()
        Panel4.Visible = False
    End Sub

    Function Load_Gride_Items_Movement()
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
        Dim _Date As Date
        Dim X As Integer
        Dim _coloumCount As Integer
        Dim Value As Double
        Dim _STSting As String
        Dim _week As Integer
        Dim _BALANCE As Double

        Try

            UltraGrid2.DisplayLayout.Bands(0).Groups.Clear()
            UltraGrid2.DisplayLayout.Bands(0).Columns.Dispose()

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
            agroup1 = UltraGrid2.DisplayLayout.Bands(0).Groups.Add("")


            agroup1.Width = 110
            Dim dt As DataTable = New DataTable()
            ' dt.Columns.Add("ID", GetType(Integer))
            Dim colWork As New DataColumn("Date", GetType(String))
            dt.Columns.Add(colWork)
            colWork.ReadOnly = True
            'colWork = New DataColumn("Item Name", GetType(String))
            'colWork.MaxLength = 250
            'dt.Columns.Add(colWork)
            'colWork.ReadOnly = True

            'dt.Columns.Add("##", GetType(String))
            ' dt.Columns.Add("Shade", GetType(String))
            I = 0
            vcWhere = "select S01Date from S01Stock_Balance where S01Status='A' and S01Item_Code='" & Trim(cboItem.Text) & "' group by S01Item_Code,S01Date"
            M01 = DBEngin.ExecuteDataset(con, Nothing, vcWhere)
            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                dt.Rows.Add(Year(M01.Tables(0).Rows(I)("S01Date")) & "/" & Month(M01.Tables(0).Rows(I)("S01Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(I)("S01Date")))
                I = I + 1
            Next

            Me.UltraGrid2.SetDataBinding(dt, Nothing)
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns(0).Group = agroup1
            'Me.UltraGrid2.DisplayLayout.Bands(0).Columns(1).Group = agroup1
            ' Me.dg_YDP_Projection.DisplayLayout.Bands(0).Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            '  Me.UltraGrid2.DisplayLayout.Bands(0).Columns(1).Width = 260
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns(0).Width = 70
            Dim _Group As String
            'agroup2.Key = ""
            'agroup3.Key = ""
            'agroup4.Key = ""
            '' agroup5.Key = ""

            'I = 0
            ''T01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            'vcWhere = "T15Sales_Order='" & strSales_Order & "' and T15Line_Item=" & txtLine_Item.Text & ""
            'M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TPR"), New SqlParameter("@vcWhereClause1", vcWhere))
            '_week = M01.Tables(0).Rows.Count
            'For Each DTRow3 As DataRow In M01.Tables(0).Rows
            '    _Group = "Group" & I + 1
            '    If I = 0 Then
            '        'agroup2.Key = ""
            agroup2 = UltraGrid2.DisplayLayout.Bands(0).Groups.Add("Group1")

            agroup2.Header.Caption = "Item Movement"
            agroup2.Width = 220

            Me.UltraGrid2.DisplayLayout.Bands(0).Columns.Add("O/B", "O/B on " & Year(txtC1.Text) & "/" & Month(txtC1.Text) & "/" & Microsoft.VisualBasic.Day(txtC1.Text))
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns("O/B").Group = agroup2
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns("O/B").Width = 110
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns("O/B").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

            Me.UltraGrid2.DisplayLayout.Bands(0).Columns.Add("Item Issue", "Item Issue")
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns("Item Issue").Group = agroup2
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns("Item Issue").Width = 70
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns("Item Issue").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

            Me.UltraGrid2.DisplayLayout.Bands(0).Columns.Add("Direct Sales", "Direct Sales")
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns("Direct Sales").Group = agroup2
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns("Direct Sales").Width = 70
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns("Direct Sales").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

            Me.UltraGrid2.DisplayLayout.Bands(0).Columns.Add("GRN", "GRN")
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns("GRN").Group = agroup2
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns("GRN").Width = 70
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns("GRN").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center


            Me.UltraGrid2.DisplayLayout.Bands(0).Columns.Add("Sup_Return", "Sup_Return")
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns("Sup_Return").Group = agroup2
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns("Sup_Return").Width = 70
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns("Sup_Return").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

            Me.UltraGrid2.DisplayLayout.Bands(0).Columns.Add("Wastage", "Wastage")
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns("Wastage").Group = agroup2
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns("Wastage").Width = 70
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns("Wastage").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

            Me.UltraGrid2.DisplayLayout.Bands(0).Columns.Add("Balance", "Balance")
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns("Balance").Group = agroup2
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns("Balance").Width = 90
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns("Balance").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            I = 0

            For Each uRow As UltraGridRow In UltraGrid2.Rows
                _BALANCE = 0
                'OPANING BALANACE
                Sql = "SELECT SUM(S01Qty) AS S01Qty FROM S01Stock_Balance WHERE S01Item_Code='" & Trim(cboItem.Text) & "' AND S01Date<'" & UltraGrid2.Rows(I).Cells(0).Value & "' AND S01Status='A' GROUP BY S01Item_Code"
                M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(M01) Then
                    UltraGrid2.Rows(I).Cells(1).Value = M01.Tables(0).Rows(0)("S01Qty")
                    _BALANCE = M01.Tables(0).Rows(0)("S01Qty")
                End If
                '===========================================================================
                'ITEM ISSUE
                Sql = "SELECT SUM(S01Qty) AS S01Qty FROM S01Stock_Balance WHERE S01Item_Code='" & Trim(cboItem.Text) & "' AND S01Date='" & UltraGrid2.Rows(I).Cells(0).Value & "' AND S01Status='A'  and S01Tr_Type='ISSUE_ITEM' GROUP BY S01Item_Code"
                M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(M01) Then
                    UltraGrid2.Rows(I).Cells(2).Value = M01.Tables(0).Rows(0)("S01Qty")
                    _BALANCE = _BALANCE + CDbl(M01.Tables(0).Rows(0)("S01Qty"))
                End If
                '==========================================================================
                'DIRECT SALES
                Sql = "SELECT SUM(S01Qty) AS S01Qty FROM S01Stock_Balance WHERE S01Item_Code='" & Trim(cboItem.Text) & "' AND S01Date='" & UltraGrid2.Rows(I).Cells(0).Value & "' AND S01Status='A'  and S01Tr_Type='DIRECT_SALES' GROUP BY S01Item_Code"
                M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(M01) Then
                    UltraGrid2.Rows(I).Cells(3).Value = M01.Tables(0).Rows(0)("S01Qty")
                    _BALANCE = _BALANCE + CDbl(M01.Tables(0).Rows(0)("S01Qty"))
                End If
                '=========================================================================
                'GRN
                Sql = "SELECT SUM(S01Qty) AS S01Qty FROM S01Stock_Balance WHERE S01Item_Code='" & Trim(cboItem.Text) & "' AND S01Date='" & UltraGrid2.Rows(I).Cells(0).Value & "' AND S01Status='A'  and S01Tr_Type='GRN' GROUP BY S01Item_Code"
                M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(M01) Then
                    UltraGrid2.Rows(I).Cells(4).Value = M01.Tables(0).Rows(0)("S01Qty")
                    _BALANCE = _BALANCE + CDbl(M01.Tables(0).Rows(0)("S01Qty"))
                End If
                '=========================================================================
                'SUPPLIER RETURN
                Sql = "SELECT SUM(S01Qty) AS S01Qty FROM S01Stock_Balance WHERE S01Item_Code='" & Trim(cboItem.Text) & "' AND S01Date='" & UltraGrid2.Rows(I).Cells(0).Value & "' AND S01Status='A'  and S01Tr_Type='SP_RETURN' GROUP BY S01Item_Code"
                M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(M01) Then
                    UltraGrid2.Rows(I).Cells(5).Value = M01.Tables(0).Rows(0)("S01Qty")
                    _BALANCE = _BALANCE + CDbl(M01.Tables(0).Rows(0)("S01Qty"))
                End If
                '========================================================================
                'WASTAGE
                Sql = "SELECT SUM(S01Qty) AS S01Qty FROM S01Stock_Balance WHERE S01Item_Code='" & Trim(cboItem.Text) & "' AND S01Date='" & UltraGrid2.Rows(I).Cells(0).Value & "' AND S01Status='A'  and S01Tr_Type='WST' GROUP BY S01Item_Code"
                M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(M01) Then
                    UltraGrid2.Rows(I).Cells(6).Value = M01.Tables(0).Rows(0)("S01Qty")
                    _BALANCE = _BALANCE + CDbl(M01.Tables(0).Rows(0)("S01Qty"))
                End If
                UltraGrid2.Rows(I).Cells(7).Value = _BALANCE
                I = I + 1
            Next


            con.ClearAllPools()
            con.close()

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
                con.close()

            End If
        End Try
    End Function
End Class