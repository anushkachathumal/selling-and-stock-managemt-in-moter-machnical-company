Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
'Imports Infragistics.Win.UltraWinGrid.RowLayoutStyle.GroupLayout
'Imports Infragistics.Win.UltraWinToolTip
'Imports Infragistics.Win.FormattedLinkLabel
'Imports Infragistics.Win.FormattedLinkLabel
'Imports Infragistics.Win.Misc
'Imports System.Diagnostics
'Imports Microsoft.Office.Interop.Excel
Imports System.Globalization

Public Class frmKnitting_Plan_WithTab
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim c_dataCustomer1 As DataTable
    Dim _CountryCode As String
    Dim _UnitCode As Integer
    Dim c_dataCustomer2 As System.Data.DataTable
    Dim c_dataCustomer3 As System.Data.DataTable
    Dim c_dataCustomer4 As System.Data.DataTable
    Dim c_dataCustomer5 As System.Data.DataTable
    Dim c_dataCustomer6 As System.Data.DataTable
    Dim c_dataCustomer1_YB As System.Data.DataTable
    Dim c_dataCustomer2_YB As System.Data.DataTable
    Dim c_dataCustomer1_KNT As System.Data.DataTable
    Dim c_dataCustomer2_KNT As System.Data.DataTable
    Dim c_dataCustomer1_Dye As System.Data.DataTable
    Dim c_dataCustomer2_Dye As System.Data.DataTable
    Dim c_dataCustomer3_Dye As System.Data.DataTable
    Dim c_dataCustomer4_Dye As System.Data.DataTable
    Dim c_dataCustomer_DyeYarn As System.Data.DataTable
    Dim c_dataCustomer_Yarn As System.Data.DataTable

    Dim c_dataCustomer_dgMC As System.Data.DataTable

    Dim Dye_Cal_Status As Boolean
    Dim c_dataCustomer1_Delivary As System.Data.DataTable

    Dim c_dataCustomerSTC As System.Data.DataTable
    Dim _DyeQuality As String
    Dim _Rowindex As Integer
    Dim _YarnQty As Double
    Dim _DyeYear As String
    Dim _DyeMonth As String
    Dim _Projection_Code As Integer
    Dim _Dye_LineItem As String
    Dim _Dye_Noof_Lane As Integer
    Dim _grgStatus As Boolean
    Dim _base30class As String
    Dim _LineItem As String

    Function Load_Gride_YDPProjection()
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
        Try

            dg_YDP_Projection.DisplayLayout.Bands(0).Groups.Clear()
            dg_YDP_Projection.DisplayLayout.Bands(0).Columns.Dispose()

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
            agroup1 = dg_YDP_Projection.DisplayLayout.Bands(0).Groups.Add("")
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
            colWork = New DataColumn("Shade", GetType(String))
            colWork.MaxLength = 250
            dt.Columns.Add(colWork)
            colWork.ReadOnly = True

            'dt.Columns.Add("##", GetType(String))
            ' dt.Columns.Add("Shade", GetType(String))
            I = 0
            vcWhere = "T15Sales_Order='" & strSales_Order & "' and T15Line_Item=" & txtLine_Item.Text & ""
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "T15"), New SqlParameter("@vcWhereClause1", vcWhere))
            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                dt.Rows.Add(M01.Tables(0).Rows(I)("T15Code"), UCase(M01.Tables(0).Rows(I)("T15Shade")))
                I = I + 1
            Next

            Me.dg_YDP_Projection.SetDataBinding(dt, Nothing)
            Me.dg_YDP_Projection.DisplayLayout.Bands(0).Columns(0).Group = agroup1
            Me.dg_YDP_Projection.DisplayLayout.Bands(0).Columns(1).Group = agroup1
            Me.dg_YDP_Projection.DisplayLayout.Bands(0).Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            Me.dg_YDP_Projection.DisplayLayout.Bands(0).Columns(0).Width = 180
            Me.dg_YDP_Projection.DisplayLayout.Bands(0).Columns(1).Width = 50
            Dim _Group As String
            'agroup2.Key = ""
            'agroup3.Key = ""
            'agroup4.Key = ""
            '' agroup5.Key = ""

            I = 0
            'T01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            vcWhere = "T15Sales_Order='" & strSales_Order & "' and T15Line_Item=" & txtLine_Item.Text & ""
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TPR"), New SqlParameter("@vcWhereClause1", vcWhere))
            _week = M01.Tables(0).Rows.Count
            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                _Group = "Group" & I + 1
                If I = 0 Then
                    'agroup2.Key = ""
                    agroup2 = dg_YDP_Projection.DisplayLayout.Bands(0).Groups.Add("Group1")

                    agroup2.Header.Caption = MonthName(Trim(M01.Tables(0).Rows(I)("T15Month")))
                    agroup2.Width = 220
                    Me.dg_YDP_Projection.DisplayLayout.Bands(0).Columns.Add("ProjectColumn", "Projection")
                    Me.dg_YDP_Projection.DisplayLayout.Bands(0).Columns("ProjectColumn").Group = agroup2
                    Me.dg_YDP_Projection.DisplayLayout.Bands(0).Columns("ProjectColumn").Width = 70
                    Me.dg_YDP_Projection.DisplayLayout.Bands(0).Columns("ProjectColumn").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                ElseIf I = 1 Then
                    agroup3 = dg_YDP_Projection.DisplayLayout.Bands(0).Groups.Add("Group2")
                    agroup3.Header.Caption = MonthName(Trim(M01.Tables(0).Rows(I)("T15Month")))
                    agroup3.Width = 220
                    Me.dg_YDP_Projection.DisplayLayout.Bands(0).Columns.Add("ProjectColumn1", "Projection")
                    Me.dg_YDP_Projection.DisplayLayout.Bands(0).Columns("ProjectColumn1").Group = agroup3
                    Me.dg_YDP_Projection.DisplayLayout.Bands(0).Columns("ProjectColumn1").Width = 70
                    Me.dg_YDP_Projection.DisplayLayout.Bands(0).Columns("ProjectColumn1").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                ElseIf I = 2 Then
                    agroup4 = dg_YDP_Projection.DisplayLayout.Bands(0).Groups.Add("Group3")
                    agroup4.Header.Caption = MonthName(Trim(M01.Tables(0).Rows(I)("T15Month")))
                    agroup4.Width = 220
                    Me.dg_YDP_Projection.DisplayLayout.Bands(0).Columns.Add("ProjectColumn2", "Projection")
                    Me.dg_YDP_Projection.DisplayLayout.Bands(0).Columns("ProjectColumn2").Group = agroup4
                    Me.dg_YDP_Projection.DisplayLayout.Bands(0).Columns("ProjectColumn2").Width = 70
                    Me.dg_YDP_Projection.DisplayLayout.Bands(0).Columns("ProjectColumn2").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                ElseIf I = 3 Then
                    agroup5 = dg_YDP_Projection.DisplayLayout.Bands(0).Groups.Add("Group4")
                    agroup5.Header.Caption = MonthName(Trim(M01.Tables(0).Rows(I)("T15Month")))
                    agroup5.Width = 220
                    Me.dg_YDP_Projection.DisplayLayout.Bands(0).Columns.Add("ProjectColumn3", "Projection")
                    Me.dg_YDP_Projection.DisplayLayout.Bands(0).Columns("ProjectColumn3").Group = agroup5
                    Me.dg_YDP_Projection.DisplayLayout.Bands(0).Columns("ProjectColumn3").Width = 70
                    Me.dg_YDP_Projection.DisplayLayout.Bands(0).Columns("ProjectColumn3").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                End If
                I = I + 1
            Next

            _coloumCount = 2
            I = 0
            vcWhere = "T15Sales_Order='" & strSales_Order & "' and T15Line_Item=" & txtLine_Item.Text & ""
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TPR"), New SqlParameter("@vcWhereClause1", vcWhere))
            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                X = 0
                _coloumCount = 2
                For Each uRow As UltraGridRow In dg_YDP_Projection.Rows
                    vcWhere = "T15Sales_Order='" & strSales_Order & "' and T15Line_Item=" & txtLine_Item.Text & " and T15Code='" & Trim(dg_YDP_Projection.Rows(X).Cells(0).Text) & "' and T15Shade='" & Trim(dg_YDP_Projection.Rows(X).Cells(1).Text) & "'"
                    dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "T15"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(dsUser) Then
                        Value = dsUser.Tables(0).Rows(0)("T15Qty")
                        _STSting = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        _STSting = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                        dg_YDP_Projection.Rows(X).Cells(_coloumCount).Value = _STSting
                        _coloumCount = _coloumCount + 1
                    Else
                        _coloumCount = _coloumCount + 1
                    End If
                    X = X + 1
                Next
            Next

            '=====================================================================
            'ALLOCATE PROJECTION
            '  _coloumCount = 2
            I = 0
            vcWhere = "T15Sales_Order='" & strSales_Order & "' and T15Line_Item=" & txtLine_Item.Text & ""
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "T15"), New SqlParameter("@vcWhereClause1", vcWhere))
            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                X = 8
                ' _coloumCount = 2
                vcWhere = "CODE='" & M01.Tables(0).Rows(I)("T15CODE") & "' and M43PRODUCT_MONTH=" & M01.Tables(0).Rows(I)("T15Month") & " and M43YEAR='" & M01.Tables(0).Rows(I)("T15Year") & "'"
                dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "VPR"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(dsUser) Then
                    Value = dsUser.Tables(0).Rows(0)("Qty") + Value
                    '    _STSting = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    '    _STSting = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                    '    dg_YDP_Projection.Rows(X).Cells(_coloumCount).Value = _STSting
                    '    _coloumCount = _coloumCount + 1
                    'Else
                    '_coloumCount = _coloumCount + 1
                End If
                I = I + 1
            Next

            X = dg1_YDP.DisplayLayout.Bands(0).Columns.Count - 1
            Value = Value / X
            _coloumCount = 1
            For I = 1 To X
                _STSting = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _STSting = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                dg1_YDP.Rows(7).Cells(_coloumCount).Value = _STSting
                _coloumCount = _coloumCount + 1
            Next

            I = 0
            Value = 0
            vcWhere = "T15Sales_Order='" & strSales_Order & "' and T15Line_Item=" & txtLine_Item.Text & ""
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "T15"), New SqlParameter("@vcWhereClause1", vcWhere))
            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                X = 8
                ' _coloumCount = 2
                vcWhere = "T15CODE='" & M01.Tables(0).Rows(I)("T15CODE") & "' and T15Month=" & M01.Tables(0).Rows(I)("T15Month") & " and T15Year='" & M01.Tables(0).Rows(I)("T15Year") & "'"
                dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TPJQ"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(dsUser) Then
                    Value = dsUser.Tables(0).Rows(0)("Qty") + Value
                    '    _STSting = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    '    _STSting = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                    '    dg_YDP_Projection.Rows(X).Cells(_coloumCount).Value = _STSting
                    '    _coloumCount = _coloumCount + 1
                    'Else
                    '_coloumCount = _coloumCount + 1
                End If
                I = I + 1
            Next

            X = dg1_YDP.DisplayLayout.Bands(0).Columns.Count - 1
            Value = Value / X
            _coloumCount = 1
            For I = 1 To X
                _STSting = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _STSting = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                dg1_YDP.Rows(8).Cells(_coloumCount).Value = _STSting
                _coloumCount = _coloumCount + 1
            Next

            X = dg1_YDP.DisplayLayout.Bands(0).Columns.Count - 1
            Value = Value / X
            _coloumCount = 1
            For I = 1 To X
                Value = dg1_YDP.Rows(7).Cells(_coloumCount).Value - dg1_YDP.Rows(8).Cells(_coloumCount).Value
                _STSting = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _STSting = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                dg1_YDP.Rows(9).Cells(_coloumCount).Value = _STSting
                _coloumCount = _coloumCount + 1
            Next


            con.close()
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
                con.close()

            End If
        End Try
    End Function

    Function Load_Gride_KNTProjection()
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
        Dim _CodeDes As String

        Try

            dg_Knt_Pojection.DisplayLayout.Bands(0).Groups.Clear()
            dg_Knt_Pojection.DisplayLayout.Bands(0).Columns.Dispose()

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
            agroup1 = dg_Knt_Pojection.DisplayLayout.Bands(0).Groups.Add("")
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
            colWork = New DataColumn("Shade", GetType(String))
            colWork.MaxLength = 250
            dt.Columns.Add(colWork)
            colWork.ReadOnly = True

            'dt.Columns.Add("##", GetType(String))
            ' dt.Columns.Add("Shade", GetType(String))
            I = 0
            vcWhere = "T15Sales_Order='" & strSales_Order & "' and T15Line_Item=" & txtLine_Item.Text & ""
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "T15"), New SqlParameter("@vcWhereClause1", vcWhere))
            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                dt.Rows.Add(M01.Tables(0).Rows(I)("T15Code"), UCase(M01.Tables(0).Rows(I)("T15Shade")))
                I = I + 1
            Next

            Me.dg_Knt_Pojection.SetDataBinding(dt, Nothing)
            Me.dg_Knt_Pojection.DisplayLayout.Bands(0).Columns(0).Group = agroup1
            Me.dg_Knt_Pojection.DisplayLayout.Bands(0).Columns(1).Group = agroup1
            Me.dg_Knt_Pojection.DisplayLayout.Bands(0).Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            Me.dg_Knt_Pojection.DisplayLayout.Bands(0).Columns(0).Width = 180
            Me.dg_Knt_Pojection.DisplayLayout.Bands(0).Columns(1).Width = 50
            Dim _Group As String
            'agroup2.Key = ""
            'agroup3.Key = ""
            'agroup4.Key = ""
            '' agroup5.Key = ""

            I = 0
            'T01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            vcWhere = "T15Sales_Order='" & strSales_Order & "' and T15Line_Item=" & txtLine_Item.Text & ""
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TPR"), New SqlParameter("@vcWhereClause1", vcWhere))
            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                _Group = "Group" & I + 1
                If I = 0 Then
                    'agroup2.Key = ""
                    agroup2 = dg_Knt_Pojection.DisplayLayout.Bands(0).Groups.Add("Group1")

                    agroup2.Header.Caption = MonthName(Trim(M01.Tables(0).Rows(I)("T15Month")))
                    agroup2.Width = 220
                    Me.dg_Knt_Pojection.DisplayLayout.Bands(0).Columns.Add("ProjectColumn", "Projection")
                    Me.dg_Knt_Pojection.DisplayLayout.Bands(0).Columns("ProjectColumn").Group = agroup2
                    Me.dg_Knt_Pojection.DisplayLayout.Bands(0).Columns("ProjectColumn").Width = 70
                    Me.dg_Knt_Pojection.DisplayLayout.Bands(0).Columns("ProjectColumn").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                ElseIf I = 1 Then
                    agroup3 = dg_Knt_Pojection.DisplayLayout.Bands(0).Groups.Add("Group2")
                    agroup3.Header.Caption = MonthName(Trim(M01.Tables(0).Rows(I)("T15Month")))
                    agroup3.Width = 220
                    Me.dg_Knt_Pojection.DisplayLayout.Bands(0).Columns.Add("ProjectColumn1", "Projection")
                    Me.dg_Knt_Pojection.DisplayLayout.Bands(0).Columns("ProjectColumn1").Group = agroup3
                    Me.dg_Knt_Pojection.DisplayLayout.Bands(0).Columns("ProjectColumn1").Width = 70
                    Me.dg_Knt_Pojection.DisplayLayout.Bands(0).Columns("ProjectColumn1").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                ElseIf I = 2 Then
                    agroup4 = dg_Knt_Pojection.DisplayLayout.Bands(0).Groups.Add("Group3")
                    agroup4.Header.Caption = MonthName(Trim(M01.Tables(0).Rows(I)("T15Month")))
                    agroup4.Width = 220
                    Me.dg_Knt_Pojection.DisplayLayout.Bands(0).Columns.Add("ProjectColumn2", "Projection")
                    Me.dg_Knt_Pojection.DisplayLayout.Bands(0).Columns("ProjectColumn2").Group = agroup4
                    Me.dg_Knt_Pojection.DisplayLayout.Bands(0).Columns("ProjectColumn2").Width = 70
                    Me.dg_Knt_Pojection.DisplayLayout.Bands(0).Columns("ProjectColumn2").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                ElseIf I = 3 Then
                    agroup5 = dg_Knt_Pojection.DisplayLayout.Bands(0).Groups.Add("Group4")
                    agroup5.Header.Caption = MonthName(Trim(M01.Tables(0).Rows(I)("T15Month")))
                    agroup5.Width = 220
                    Me.dg_Knt_Pojection.DisplayLayout.Bands(0).Columns.Add("ProjectColumn3", "Projection")
                    Me.dg_Knt_Pojection.DisplayLayout.Bands(0).Columns("ProjectColumn3").Group = agroup5
                    Me.dg_Knt_Pojection.DisplayLayout.Bands(0).Columns("ProjectColumn3").Width = 70
                    Me.dg_Knt_Pojection.DisplayLayout.Bands(0).Columns("ProjectColumn3").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                End If
                I = I + 1
            Next

            _coloumCount = 2
            I = 0
            vcWhere = "T15Sales_Order='" & strSales_Order & "' and T15Line_Item=" & txtLine_Item.Text & ""
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TPR"), New SqlParameter("@vcWhereClause1", vcWhere))
            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                X = 0
                ' _coloumCount = 2
                For Each uRow As UltraGridRow In dg_Knt_Pojection.Rows
                    '**---------------------Previous Code---------------***

                    '======vcWhere = "T15Sales_Order='" & strSales_Order & "' and T15Line_Item=" & txtLine_Item.Text & " and T15Code='" & Trim(dg_Knt_Pojection.Rows(X).Cells(0).Text) & "' and T15Shade='" & Trim(dg_Knt_Pojection.Rows(X).Cells(1).Text) & "'"
                    '======dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "T15"), New SqlParameter("@vcWhereClause1", vcWhere))

                    'New Code redeveloped by Suranga Wijesinghe on 2016.6.14
                    'Request by Amila Priyankara
                    '**---------------------NEW CODE---------------------***
                    _CodeDes = Trim(dg_Knt_Pojection.Rows(X).Cells(0).Text) & " | " & Trim(dg_Knt_Pojection.Rows(X).Cells(1).Text)
                    vcWhere = "Code='" & _CodeDes & "' and M43Year='" & M01.Tables(0).Rows(I)("T15Year") & "' and M43Product_Month='" & M01.Tables(0).Rows(I)("T15Month") & "'"
                    dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "VM43"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(dsUser) Then
                        Value = dsUser.Tables(0).Rows(0)("Qty")
                        _STSting = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        _STSting = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                        dg_Knt_Pojection.Rows(X).Cells(_coloumCount).Value = _STSting
                        ' _coloumCount = _coloumCount + 1
                    Else
                        ' _coloumCount = _coloumCount + 1
                    End If
                    X = X + 1
                Next
                _coloumCount = _coloumCount + 1
                I = I + 1
            Next
            con.close()
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
                con.close()

            End If
        End Try
    End Function

    Function Load_Grid_Delivary_Plan()
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
        Dim _CodeDes As String
        Dim _ST As String

        Try
            dgDel_Plan.DisplayLayout.Bands(0).Groups.Clear()
            dgDel_Plan.DisplayLayout.Bands(0).Columns.Dispose()

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
            agroup1 = dgDel_Plan.DisplayLayout.Bands(0).Groups.Add("")


            agroup1.Width = 110
            Dim dt As DataTable = New DataTable()
            ' dt.Columns.Add("ID", GetType(Integer))
            Dim colWork As New DataColumn("Line Item", GetType(String))
            dt.Columns.Add(colWork)
            colWork.ReadOnly = True
            colWork = New DataColumn("Trim", GetType(String))
            colWork.MaxLength = 250
            dt.Columns.Add(colWork)
            colWork.ReadOnly = True

            '  dt.Rows.Add(M01.Tables(0).Rows(I)("T15Code"), UCase(M01.Tables(0).Rows(I)("T15Shade")))
            dt.Rows.Add(_LineItem)

            Me.dgDel_Plan.SetDataBinding(dt, Nothing)
            Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(0).Group = agroup1
            Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(1).Group = agroup1
            Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(0).Width = 70
            Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(1).Width = 70

            Dim _Count As Integer
            Dim _WeeNo As String
            Dim _WeekNo As String
            Dim _cOLUM As Integer

            _cOLUM = 2
            vcWhere = "tmpSales_Order='" & strSales_Order & "' and tmpLine_Item in ('" & _LineItem & "')"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "KDLP"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                agroup3 = dgDel_Plan.DisplayLayout.Bands(0).Groups.Add("Knt")
                agroup3.Header.Caption = "Knitting"
                _Count = M01.Tables(0).Rows.Count


                'KNITTIG
                I = 0
                For Each DTRow3 As DataRow In M01.Tables(0).Rows
                    _WeeNo = "WK-" & M01.Tables(0).Rows(I)("tmpWeek_No")
                    'Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(2).Width = 70 * _Count
                    Me.dgDel_Plan.DisplayLayout.Bands(0).Columns.Add(_WeeNo, _WeeNo)
                    Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(_WeeNo).Group = agroup3
                    Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(_WeeNo).Width = 90
                    Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(_WeeNo).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                    Value = M01.Tables(0).Rows(I)("qTY")
                    _STSting = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    _STSting = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                    dgDel_Plan.Rows(0).Cells(_cOLUM).Value = _STSting
                    ' dgDel_Plan.Rows(0).Cells(_cOLUM).Appearance.TextHAlign = Infragistics.Win.HAlign.Center
                    I = I + 1
                    _cOLUM = _cOLUM + 1
                Next

            End If
            '============================================================================================
            vcWhere = "T18Sales_Order='" & strSales_Order & "' and T18Line_Item in ('" & _LineItem & "')"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "DELZ"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                agroup4 = dgDel_Plan.DisplayLayout.Bands(0).Groups.Add("dye")
                agroup4.Header.Caption = "Dyeing"
                _Count = M01.Tables(0).Rows.Count

                'agroup5 = dgDel_Plan.DisplayLayout.Bands(0).Groups.Add("DEL")
                'agroup5.Header.Caption = "Delivery"
                'DYEING
                I = 0
                For Each DTRow3 As DataRow In M01.Tables(0).Rows
                    _WeeNo = "Wk_" & M01.Tables(0).Rows(I)("T18WeekNo")
                    'Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(2).Width = 70 * _Count
                    Me.dgDel_Plan.DisplayLayout.Bands(0).Columns.Add(_WeeNo, _WeeNo)
                    Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(_WeeNo).Group = agroup4
                    Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(_WeeNo).Width = 90
                    Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(_WeeNo).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                    Value = M01.Tables(0).Rows(I)("qTY")
                    If Value > 0 Then
                        _STSting = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        _STSting = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                        dgDel_Plan.Rows(0).Cells(_cOLUM).Value = _STSting
                    ElseIf Trim(M01.Tables(0).Rows(I)("SUB")) = "Y" Then
                        dgDel_Plan.Rows(0).Cells(_cOLUM).Value = "Sub"
                        dgDel_Plan.Rows(0).Cells(_cOLUM).Appearance.TextHAlign = Infragistics.Win.HAlign.Center
                    ElseIf Trim(M01.Tables(0).Rows(I)("App")) = "Y" Then
                        dgDel_Plan.Rows(0).Cells(_cOLUM).Value = "App"
                        dgDel_Plan.Rows(0).Cells(_cOLUM).Appearance.TextHAlign = Infragistics.Win.HAlign.Center
                    End If
                    _cOLUM = _cOLUM + 1

                    '_WeeNo = "WK_" & (M01.Tables(0).Rows(I)("T18WeekNo") + 1)
                    ''Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(2).Width = 70 * _Count
                    'Me.dgDel_Plan.DisplayLayout.Bands(0).Columns.Add(_WeeNo, _WeeNo)
                    'Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(_WeeNo).Group = agroup5
                    'Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(_WeeNo).Width = 90
                    'Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(_WeeNo).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                    I = I + 1
                Next

                
            End If

            vcWhere = "select * from T18Delivary_Plane where T18Sales_Order='" & strSales_Order & "' and T18Line_Item in ('" & _LineItem & "') and T18App='Y' order by T18Year,T18WeekNo desc "
            dsUser = DBEngin.ExecuteDataset(con, Nothing, vcWhere)
            If isValidDataset(dsUser) Then
                Dim _Wk As Integer
                _Wk = dsUser.Tables(0).Rows(0)("T18WeekNo")

                vcWhere = "T18Sales_Order='" & strSales_Order & "' and T18Line_Item in ('" & _LineItem & "')"
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "DELZ"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    agroup5 = dgDel_Plan.DisplayLayout.Bands(0).Groups.Add("DEL")
                    agroup5.Header.Caption = "Delivery"
                    'DYEING
                    I = 0
                    For Each DTRow3 As DataRow In M01.Tables(0).Rows
                        '_WeeNo = "WK_" & M01.Tables(0).Rows(I)("T18WeekNo")
                        ''Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(2).Width = 70 * _Count
                        'Me.dgDel_Plan.DisplayLayout.Bands(0).Columns.Add(_WeeNo, _WeeNo)
                        'Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(_WeeNo).Group = agroup5
                        'Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(_WeeNo).Width = 90
                        'Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(_WeeNo).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                        Dim dfi As DateTimeFormatInfo = DateTimeFormatInfo.CurrentInfo
                        Dim date1 As Date = "1/1/" & Year(Today)
                        Dim cal As Calendar = dfi.Calendar

                        Value = M01.Tables(0).Rows(I)("qTY")
                        If Value > 0 Then
                            _WeekNo = (M01.Tables(0).Rows(I)("T18WeekNo") + 1)
                            _WeeNo = "WK-" & (_Wk + 1)
                            'Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(2).Width = 70 * _Count
                            Me.dgDel_Plan.DisplayLayout.Bands(0).Columns.Add(_WeekNo, _WeeNo)
                            Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(_WeekNo).Group = agroup5
                            Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(_WeekNo).Width = 90
                            Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(_WeekNo).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                            'Value = M01.Tables(0).Rows(I)("qTY")
                            _STSting = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                            _STSting = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                            dgDel_Plan.Rows(0).Cells(_cOLUM).Value = _STSting
                            _cOLUM = _cOLUM + 1
                            _Wk = _Wk + 1

                            If cal.GetWeekOfYear(date1, dfi.CalendarWeekRule, dfi.FirstDayOfWeek) >= _Wk Then
                                _Wk = 0
                            End If
                        End If
                        I = I + 1
                    Next
                End If
            Else

                vcWhere = "T18Sales_Order='" & strSales_Order & "' and T18Line_Item in ('" & _LineItem & "')"
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "DELZ"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    agroup5 = dgDel_Plan.DisplayLayout.Bands(0).Groups.Add("DEL")
                    agroup5.Header.Caption = "Delivery"
                    'DYEING
                    I = 0
                    For Each DTRow3 As DataRow In M01.Tables(0).Rows
                        '_WeeNo = "WK_" & M01.Tables(0).Rows(I)("T18WeekNo")
                        ''Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(2).Width = 70 * _Count
                        'Me.dgDel_Plan.DisplayLayout.Bands(0).Columns.Add(_WeeNo, _WeeNo)
                        'Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(_WeeNo).Group = agroup5
                        'Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(_WeeNo).Width = 90
                        'Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(_WeeNo).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                        Value = M01.Tables(0).Rows(I)("qTY")
                        If Value > 0 Then
                            _WeekNo = (M01.Tables(0).Rows(I)("T18WeekNo") + 1)
                            _WeeNo = "WK-" & (M01.Tables(0).Rows(I)("T18WeekNo") + 1)
                            'Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(2).Width = 70 * _Count
                            Me.dgDel_Plan.DisplayLayout.Bands(0).Columns.Add(_WeekNo, _WeeNo)
                            Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(_WeekNo).Group = agroup5
                            Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(_WeekNo).Width = 90
                            Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(_WeekNo).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                            'Value = M01.Tables(0).Rows(I)("qTY")
                            _STSting = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                            _STSting = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                            dgDel_Plan.Rows(0).Cells(_cOLUM).Value = _STSting
                            _cOLUM = _cOLUM + 1
                        End If
                        I = I + 1
                    Next
                End If
            End If
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            con.close()
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)

                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
                con.close()
            End If
        End Try
    End Function

    Function Load_Grid_Delivary_Plan_SLTime()
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
        Dim _CodeDes As String
        Dim _ST As String

        Try
            dgDel_Plan.DisplayLayout.Bands(0).Groups.Clear()
            dgDel_Plan.DisplayLayout.Bands(0).Columns.Dispose()

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
            agroup1 = dgDel_Plan.DisplayLayout.Bands(0).Groups.Add("")


            agroup1.Width = 110
            Dim dt As DataTable = New DataTable()
            ' dt.Columns.Add("ID", GetType(Integer))
            Dim colWork As New DataColumn("Line Item", GetType(String))
            dt.Columns.Add(colWork)
            colWork.ReadOnly = True
            colWork = New DataColumn("Trim", GetType(String))
            colWork.MaxLength = 250
            dt.Columns.Add(colWork)
            colWork.ReadOnly = True

            '  dt.Rows.Add(M01.Tables(0).Rows(I)("T15Code"), UCase(M01.Tables(0).Rows(I)("T15Shade")))
            dt.Rows.Add(_LineItem)

            Me.dgDel_Plan.SetDataBinding(dt, Nothing)
            Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(0).Group = agroup1
            Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(1).Group = agroup1
            Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(0).Width = 70
            Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(1).Width = 70

            Dim _Count As Integer
            Dim _WeeNo As String
            Dim _cOLUM As Integer

            _cOLUM = 2
            vcWhere = "tmpSales_Order='" & strSales_Order & "' and tmpLine_Item in ('" & _LineItem & "')"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "KDLP"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                agroup3 = dgDel_Plan.DisplayLayout.Bands(0).Groups.Add("Knt")
                agroup3.Header.Caption = "Knitting"
                _Count = M01.Tables(0).Rows.Count


                'KNITTIG
                I = 0
                For Each DTRow3 As DataRow In M01.Tables(0).Rows
                    _WeeNo = "WK-" & M01.Tables(0).Rows(I)("tmpWeek_No")
                    'Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(2).Width = 70 * _Count
                    Me.dgDel_Plan.DisplayLayout.Bands(0).Columns.Add(_WeeNo, _WeeNo)
                    Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(_WeeNo).Group = agroup3
                    Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(_WeeNo).Width = 90
                    Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(_WeeNo).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                    Value = M01.Tables(0).Rows(I)("qTY")
                    _STSting = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    _STSting = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                    dgDel_Plan.Rows(0).Cells(_cOLUM).Value = _STSting
                    I = I + 1
                    _cOLUM = _cOLUM + 1
                Next

            End If
            '============================================================================================
            vcWhere = "T18Sales_Order='" & strSales_Order & "' and T18Line_Item in ('" & _LineItem & "')"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "DELZ"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                agroup4 = dgDel_Plan.DisplayLayout.Bands(0).Groups.Add("dye")
                agroup4.Header.Caption = "Dyeing"
                _Count = M01.Tables(0).Rows.Count

                'agroup5 = dgDel_Plan.DisplayLayout.Bands(0).Groups.Add("DEL")
                'agroup5.Header.Caption = "Delivery"
                'DYEING
                I = 0
                For Each DTRow3 As DataRow In M01.Tables(0).Rows
                    _WeeNo = "Wk_" & M01.Tables(0).Rows(I)("T18WeekNo")
                    'Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(2).Width = 70 * _Count
                    Me.dgDel_Plan.DisplayLayout.Bands(0).Columns.Add(_WeeNo, _WeeNo)
                    Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(_WeeNo).Group = agroup4
                    Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(_WeeNo).Width = 90
                    Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(_WeeNo).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                    Value = M01.Tables(0).Rows(I)("qTY")
                    _STSting = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    _STSting = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                    dgDel_Plan.Rows(0).Cells(_cOLUM).Value = _STSting
                    _cOLUM = _cOLUM + 1

                    '_WeeNo = "WK_" & (M01.Tables(0).Rows(I)("T18WeekNo") + 1)
                    ''Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(2).Width = 70 * _Count
                    'Me.dgDel_Plan.DisplayLayout.Bands(0).Columns.Add(_WeeNo, _WeeNo)
                    'Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(_WeeNo).Group = agroup5
                    'Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(_WeeNo).Width = 90
                    'Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(_WeeNo).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                    I = I + 1
                Next


            End If


            'vcWhere = "T18Sales_Order='" & strSales_Order & "' and T18Line_Item in ('" & _LineItem & "')"
            'M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "DELZ"), New SqlParameter("@vcWhereClause1", vcWhere))
            'If isValidDataset(M01) Then
            '    agroup5 = dgDel_Plan.DisplayLayout.Bands(0).Groups.Add("DEL")
            '    agroup5.Header.Caption = "Delivery"
            '    'DYEING
            '    I = 0
            '    For Each DTRow3 As DataRow In M01.Tables(0).Rows
            '        '_WeeNo = "WK_" & M01.Tables(0).Rows(I)("T18WeekNo")
            '        ''Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(2).Width = 70 * _Count
            '        'Me.dgDel_Plan.DisplayLayout.Bands(0).Columns.Add(_WeeNo, _WeeNo)
            '        'Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(_WeeNo).Group = agroup5
            '        'Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(_WeeNo).Width = 90
            '        'Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(_WeeNo).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right



            '        _WeeNo = "WK_" & (M01.Tables(0).Rows(I)("T18WeekNo") + 1)
            '        'Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(2).Width = 70 * _Count
            '        Me.dgDel_Plan.DisplayLayout.Bands(0).Columns.Add(_WeeNo, _WeeNo)
            '        Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(_WeeNo).Group = agroup5
            '        Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(_WeeNo).Width = 90
            '        Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(_WeeNo).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            '        Value = M01.Tables(0).Rows(I)("qTY")
            '        _STSting = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            '        _STSting = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

            '        dgDel_Plan.Rows(0).Cells(_cOLUM).Value = _STSting
            '        _cOLUM = _cOLUM + 1
            '        I = I + 1
            '    Next
            'End If

            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            con.close()
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)

                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
                con.close()
            End If
        End Try
    End Function


    Function Load_Gride_YB()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1_YB = CustomerDataClass.MakeDataTableYarn_Booking
        dg1_YB.DataSource = c_dataCustomer1_YB
        With dg1_YB
            .DisplayLayout.Bands(0).Columns(0).Width = 80
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 230
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 90
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(3).Width = 80
            .DisplayLayout.Bands(0).Columns(3).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(4).Width = 60
            .DisplayLayout.Bands(0).Columns(4).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(5).Width = 90
            .DisplayLayout.Bands(0).Columns(5).AutoEdit = False

            .DisplayLayout.Bands(0).Columns(6).Width = 90
            .DisplayLayout.Bands(0).Columns(6).AutoEdit = True

            '   .DisplayLayout.Bands(0).Columns(6).AutoEdit = False
            ' .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Left
            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
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

    Function Load_GrideStock_YB()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer2_YB = CustomerDataClass.MakeDataTablePreVious_Stock
        dg2_YB.DataSource = c_dataCustomer2_YB
        With dg2_YB
            .DisplayLayout.Bands(0).Columns(0).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 40
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(3).Width = 40
            .DisplayLayout.Bands(0).Columns(3).AutoEdit = False


            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            '  .DisplayLayout.Bands(0).Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            '.DisplayLayout.Bands(0).Columns(3).Width = 90
            '.DisplayLayout.Bands(0).Columns(4).Width = 90
            '.DisplayLayout.Bands(0).Columns(5).Width = 90
            ''  .DisplayLayout.Bands(0).Columns(6).Width = 90
            ' .DisplayLayout.Bands(0).Columns(7).Width = 90

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub

    Function Load_Parameter()
        Dim M01 As DataSet
        Dim M02 As DataSet


        Try
            Dim con = New SqlConnection()
            con = DBEngin.GetConnection(True)

            ''vcWhere = " M22Quality='" & Trim(txtQuality.Text) & "' and M22Strich_Lenth>0"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetUse_Griege_Qty", New SqlParameter("@cQryType", "PAR"))
            If isValidDataset(M01) Then
                GrgRef = M01.Tables(0).Rows(0)("P01NO")
            End If

            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                'Con.close()
            End If
        End Try
    End Function

    Function Update_Parameter()
        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Dim i As Integer

        Try
            nvcFieldList1 = "update P01PARAMETER set P01NO=P01NO +" & 1 & " where P01CODE='GP'"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)


            transaction.Commit()
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
            connection.Close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
                connection.Close()
            End If
        End Try

    End Function

    Function Update_Date()
        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Dim i As Integer
        Dim vcWhere As String
        Dim M01 As DataSet
        Dim ncQryType As String



        Try
            vcWhere = "T11Ref_No=" & Delivary_Ref & " and T11Sales_Order='" & txtSO.Text & "' and T11Line_Item=" & txtLine_Item.Text & ""
            M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "LADC"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                nvcFieldList1 = "update T11Lab_Dip_ConfDate set T11Date='" & txtDate.Text & "' where T11Ref_No=" & Delivary_Ref & " and T11Sales_Order='" & txtSO.Text & "' and T11Line_Item=" & txtLine_Item.Text & ""
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            Else
                ncQryType = "LADU"
                nvcFieldList1 = "(T11Ref_No," & "T11Sales_Order," & "T11Line_Item," & "T11Date," & "T11User) " & "values(" & Delivary_Ref & ",'" & txtSO.Text & "'," & txtLine_Item.Text & ",'" & txtDate.Text & "','" & strDisname & "')"
                up_GetSetLAB_DIP_DateConf(ncQryType, nvcFieldList1, vcWhere, connection, transaction)
            End If

            transaction.Commit()
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
            connection.Close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                '  MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
                connection.Close()
            End If
        End Try
    End Function

    Function Search_Data()
        Dim vcWhere As String
        Dim M01 As DataSet
        Dim M02 As DataSet


        Try
            Dim con = New SqlConnection()
            con = DBEngin.GetConnection(True)
            Dim TestString As String
            Dim TestArray() As String

            vcWhere = " M22Quality='" & Trim(txtQuality.Text) & "' and M22Strich_Lenth>0"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TEC"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                ' With frmGriege_Stock
                txtGauge.Text = M01.Tables(0).Rows(0)("M22Machine_Type")
                TestString = M01.Tables(0).Rows(0)("M22Machine_Type")
                TestArray = Split(TestString)
                strMC_Group = TestArray(2) & TestArray(0)
                'End With
            End If
            'COMMON QUALITY

            vcWhere = " M26Quality30='" & Trim(txtQuality.Text) & "' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "COM"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                ' With frmGriege_Stock
                txtCommon.Text = M01.Tables(0).Rows(0)("M26Quality20")
                'End With
            Else
                'With frmGriege_Stock
                txtCommon.Text = "NO"
                'End With
            End If
            '------------------------------------------------------------------
            'SUTABLE GRIGE
            vcWhere = " M14Order='" & Trim(txtRcode.Text) & "' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "GRG"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                'With frmGriege_Stock
                txtShade.Text = M01.Tables(0).Rows(0)("M14Grige")
                'End With
            Else
                'With frmGriege_Stock
                '    .txtCommon.Text = "NO"
                'End With
            End If
            '-----------------------------------------------------------------
            'PERDAY KNITTING OUTPUT
            'Coment on 02/16/2016 by Suranga Requested by Sameera Planning

            Dim Value As Double
            'vcWhere = " M22Quality='" & Trim(txtQuality.Text) & "' and RIGHT(LEFT(M22Yarn,6),2)='NE' "
            'dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPER_DAY_KNTOUTPUT", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWhere))
            'If isValidDataset(dsUser) Then
            '    '  If Microsoft.VisualBasic.Left(Trim(txtQuality.Text), 2) = "Y1" Or Microsoft.VisualBasic.Left(Trim(txtQuality.Text), 2) = "Y3" Then
            '    vcWhere = "left(M22Quality,2)in ('Y1','Y3') and M22Fabric_Type='SINGLE JERSEY' and M22Machine_Type like '%Auto Stripe%' and M22Strich_Lenth<>0 and M22Quality='" & Trim(txtQuality.Text) & "' "
            '    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPER_DAY_KNTOUTPUT", New SqlParameter("@cQryType", "RPS1"), New SqlParameter("@vcWhereClause1", vcWhere))
            '    If isValidDataset(M01) Then
            '        Value = M01.Tables(0).Rows(0)("kgH")
            '        strPer_Day = Value
            '        '   With frmGriege_Stock
            '        txtPer_Day.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            '        txtPer_Day.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            '        'End With
            '        ' Exit Function

            '    End If

            '    vcWhere = "left(M22Quality,2)in ('Y1') and M22Fabric_Type='SINGLE JERSEY' and M22Machine_Type not like '%Auto Stripe%' and M22Strich_Lenth<>0 and M22Quality='" & Trim(txtQuality.Text) & "' "
            '    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPER_DAY_KNTOUTPUT", New SqlParameter("@cQryType", "RPS2"), New SqlParameter("@vcWhereClause1", vcWhere))
            '    If isValidDataset(M01) Then
            '        Value = M01.Tables(0).Rows(0)("kgH")
            '        strPer_Day = Value
            '        'With frmGriege_Stock
            '        txtPer_Day.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            '        txtPer_Day.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            '        'End With
            '        ' Exit Function

            '    End If


            '    vcWhere = "left(M22Quality,1) not in ('Y') and M22Yarn_Cons='100' and M22Fabric_Type='SINGLE JERSEY' and M22Strich_Lenth<>0 and M22Quality='" & Trim(txtQuality.Text) & "' "
            '    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPER_DAY_KNTOUTPUT", New SqlParameter("@cQryType", "RPS2"), New SqlParameter("@vcWhereClause1", vcWhere))
            '    If isValidDataset(M01) Then
            '        Value = M01.Tables(0).Rows(0)("kgH")
            '        strPer_Day = Value
            '        'With frmGriege_Stock
            '        txtPer_Day.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            '        txtPer_Day.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            '        'End With
            '        ' Exit Function

            '    End If

            '    vcWhere = "left(M22Quality,2)  in ('Y3') and M22Yarn_Cons<'100' and M22Fabric_Type='SINGLE JERSEY' and M22Machine_Type not like '%Auto Stripe%' and M22Strich_Lenth<>0 and M22Quality='" & Trim(txtQuality.Text) & "' "
            '    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPER_DAY_KNTOUTPUT", New SqlParameter("@cQryType", "RPS3"), New SqlParameter("@vcWhereClause1", vcWhere))
            '    If isValidDataset(M01) Then
            '        Value = M01.Tables(0).Rows(0)("kgH")
            '        strPer_Day = Value
            '        'With frmGriege_Stock
            '        txtPer_Day.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            '        txtPer_Day.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            '        'End With
            '        ' Exit Function

            '    End If

            '    vcWhere = "left(M22Quality,1) NOT in ('Y') and M22Yarn_Cons<'100' and M22Fabric_Type='SINGLE JERSEY' and M22Strich_Lenth<>0 and M22Quality='" & Trim(txtQuality.Text) & "' "
            '    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPER_DAY_KNTOUTPUT", New SqlParameter("@cQryType", "RPS3"), New SqlParameter("@vcWhereClause1", vcWhere))
            '    If isValidDataset(M01) Then
            '        Value = M01.Tables(0).Rows(0)("kgH")
            '        strPer_Day = Value
            '        'With frmGriege_Stock
            '        txtPer_Day.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            '        txtPer_Day.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            '        'End With
            '        ' Exit Function

            '    End If

            '    'End If

            'End If
            ''------------------------------------------------------------------------------

            'vcWhere = " M22Quality='" & Trim(txtQuality.Text) & "' and RIGHT(LEFT(M22Yarn,6),2)='NM' "
            'dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPER_DAY_KNTOUTPUT", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWhere))
            'If isValidDataset(dsUser) Then
            '    '  If Microsoft.VisualBasic.Left(Trim(txtQuality.Text), 2) = "Y1" Or Microsoft.VisualBasic.Left(Trim(txtQuality.Text), 2) = "Y3" Then
            '    vcWhere = "left(M22Quality,2)in ('Y1','Y3') and M22Fabric_Type='SINGLE JERSEY' and M22Machine_Type like '%Auto Stripe%' and M22Strich_Lenth<>0 and M22Quality='" & Trim(txtQuality.Text) & "' "
            '    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPER_DAY_KNTOUTPUT", New SqlParameter("@cQryType", "RPN1"), New SqlParameter("@vcWhereClause1", vcWhere))
            '    If isValidDataset(M01) Then
            '        Value = M01.Tables(0).Rows(0)("kgH")
            '        strPer_Day = Value
            '        'With frmGriege_Stock
            '        txtPer_Day.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            '        txtPer_Day.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            '        'End With
            '        '   Exit Function

            '    End If

            '    vcWhere = "left(M22Quality,2)in ('Y1') and M22Fabric_Type='SINGLE JERSEY' and M22Machine_Type not like '%Auto Stripe%' and M22Strich_Lenth<>0 and M22Quality='" & Trim(txtQuality.Text) & "' "
            '    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPER_DAY_KNTOUTPUT", New SqlParameter("@cQryType", "RPN2"), New SqlParameter("@vcWhereClause1", vcWhere))
            '    If isValidDataset(M01) Then
            '        Value = M01.Tables(0).Rows(0)("kgH")
            '        strPer_Day = Value

            '        'With frmGriege_Stock
            '        txtPer_Day.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            '        txtPer_Day.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            '        'End With
            '        '   Exit Function

            '    End If


            '    vcWhere = "left(M22Quality,1) not in ('Y') and M22Yarn_Cons='100' and M22Fabric_Type='SINGLE JERSEY' and M22Strich_Lenth<>0 and M22Quality='" & Trim(txtQuality.Text) & "' "
            '    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPER_DAY_KNTOUTPUT", New SqlParameter("@cQryType", "RPN2"), New SqlParameter("@vcWhereClause1", vcWhere))
            '    If isValidDataset(M01) Then
            '        Value = M01.Tables(0).Rows(0)("kgH")
            '        strPer_Day = Value
            '        'With frmGriege_Stock
            '        txtPer_Day.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            '        txtPer_Day.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            '        'End With
            '        ' Exit Function

            '    End If

            '    vcWhere = "left(M22Quality,2)  in ('Y3') and M22Yarn_Cons<'100' and M22Fabric_Type='SINGLE JERSEY' and M22Machine_Type not like '%Auto Stripe%' and M22Strich_Lenth<>0 and M22Quality='" & Trim(txtQuality.Text) & "' "
            '    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPER_DAY_KNTOUTPUT", New SqlParameter("@cQryType", "RPN3"), New SqlParameter("@vcWhereClause1", vcWhere))
            '    If isValidDataset(M01) Then
            '        Value = M01.Tables(0).Rows(0)("kgH")
            '        strPer_Day = Value
            '        'With frmGriege_Stock
            '        txtPer_Day.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            '        txtPer_Day.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            '        'End With
            '        '  Exit Function

            '    End If

            '    vcWhere = "left(M22Quality,1) NOT in ('Y') and M22Yarn_Cons<'100' and M22Fabric_Type='SINGLE JERSEY' and M22Strich_Lenth<>0 and M22Quality='" & Trim(txtQuality.Text) & "' "
            '    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPER_DAY_KNTOUTPUT", New SqlParameter("@cQryType", "RPN3"), New SqlParameter("@vcWhereClause1", vcWhere))
            '    If isValidDataset(M01) Then
            '        Value = M01.Tables(0).Rows(0)("kgH")
            '        strPer_Day = Value
            '        'With frmGriege_Stock
            '        txtPer_Day.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            '        txtPer_Day.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            '        'End With
            '        '   Exit Function

            '    End If

            '    'End If

            'End If
            ''DT
            'vcWhere = " M22Quality='" & Trim(txtQuality.Text) & "' and RIGHT(LEFT(M22Yarn,6),2)='DT' "
            'dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPER_DAY_KNTOUTPUT", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWhere))
            'If isValidDataset(dsUser) Then
            '    '  If Microsoft.VisualBasic.Left(Trim(txtQuality.Text), 2) = "Y1" Or Microsoft.VisualBasic.Left(Trim(txtQuality.Text), 2) = "Y3" Then
            '    vcWhere = "left(M22Quality,2)in ('Y1','Y3') and M22Fabric_Type='SINGLE JERSEY' and M22Machine_Type like '%Auto Stripe%' and M22Strich_Lenth<>0 and M22Quality='" & Trim(txtQuality.Text) & "' "
            '    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPER_DAY_KNTOUTPUT", New SqlParameter("@cQryType", "RPD1"), New SqlParameter("@vcWhereClause1", vcWhere))
            '    If isValidDataset(M01) Then
            '        Value = M01.Tables(0).Rows(0)("kgH")
            '        strPer_Day = Value
            '        'With frmGriege_Stock
            '        txtPer_Day.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            '        txtPer_Day.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            '        'End With
            '        '  Exit Function

            '    End If

            '    vcWhere = "left(M22Quality,2)in ('Y1') and M22Fabric_Type='SINGLE JERSEY' and M22Machine_Type not like '%Auto Stripe%' and M22Strich_Lenth<>0 and M22Quality='" & Trim(txtQuality.Text) & "' "
            '    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPER_DAY_KNTOUTPUT", New SqlParameter("@cQryType", "RPD2"), New SqlParameter("@vcWhereClause1", vcWhere))
            '    If isValidDataset(M01) Then
            '        Value = M01.Tables(0).Rows(0)("kgH")
            '        strPer_Day = Value
            '        'With frmGriege_Stock
            '        txtPer_Day.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            '        txtPer_Day.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            '        'End With
            '        '  Exit Function

            '    End If


            '    vcWhere = "left(M22Quality,1) not in ('Y') and M22Yarn_Cons='100' and M22Fabric_Type='SINGLE JERSEY' and M22Strich_Lenth<>0 and M22Quality='" & Trim(txtQuality.Text) & "' "
            '    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPER_DAY_KNTOUTPUT", New SqlParameter("@cQryType", "RPD2"), New SqlParameter("@vcWhereClause1", vcWhere))
            '    If isValidDataset(M01) Then
            '        Value = M01.Tables(0).Rows(0)("kgH")
            '        strPer_Day = Value
            '        'With frmGriege_Stock
            '        txtPer_Day.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            '        txtPer_Day.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            '        'End With
            '        ' Exit Function

            '    End If

            '    vcWhere = "left(M22Quality,2)  in ('Y3') and M22Yarn_Cons<'100' and M22Fabric_Type='SINGLE JERSEY' and M22Machine_Type not like '%Auto Stripe%' and M22Strich_Lenth<>0 and M22Quality='" & Trim(txtQuality.Text) & "' "
            '    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPER_DAY_KNTOUTPUT", New SqlParameter("@cQryType", "RPD3"), New SqlParameter("@vcWhereClause1", vcWhere))
            '    If isValidDataset(M01) Then
            '        Value = M01.Tables(0).Rows(0)("kgH")
            '        strPer_Day = Value
            '        'With frmGriege_Stock
            '        txtPer_Day.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            '        txtPer_Day.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            '        'End With
            '        ' Exit Function

            '    End If

            '    vcWhere = "left(M22Quality,1) NOT in ('Y') and M22Yarn_Cons<'100' and M22Fabric_Type='SINGLE JERSEY' and M22Strich_Lenth<>0 and M22Quality='" & Trim(txtQuality.Text) & "' "
            '    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPER_DAY_KNTOUTPUT", New SqlParameter("@cQryType", "RPD3"), New SqlParameter("@vcWhereClause1", vcWhere))
            '    If isValidDataset(M01) Then
            '        Value = M01.Tables(0).Rows(0)("kgH")
            '        strPer_Day = Value
            '        'With frmGriege_Stock
            '        txtPer_Day.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            '        txtPer_Day.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            '        'End With
            '        '  Exit Function

            '    End If

            '    'End If

            'End If
            ''------------------------------------------------------------------------------
            ''DE
            'vcWhere = " M22Quality='" & Trim(txtQuality.Text) & "' and RIGHT(LEFT(M22Yarn,6),2)IN ('DE','D') "
            'dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPER_DAY_KNTOUTPUT", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWhere))
            'If isValidDataset(dsUser) Then
            '    '  If Microsoft.VisualBasic.Left(Trim(txtQuality.Text), 2) = "Y1" Or Microsoft.VisualBasic.Left(Trim(txtQuality.Text), 2) = "Y3" Then
            '    vcWhere = "left(M22Quality,2)in ('Y1','Y3') and M22Fabric_Type='SINGLE JERSEY' and M22Machine_Type like '%Auto Stripe%' and M22Strich_Lenth<>0 and M22Quality='" & Trim(txtQuality.Text) & "' "
            '    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPER_DAY_KNTOUTPUT", New SqlParameter("@cQryType", "RPE1"), New SqlParameter("@vcWhereClause1", vcWhere))
            '    If isValidDataset(M01) Then
            '        Value = M01.Tables(0).Rows(0)("kgH")
            '        strPer_Day = Value
            '        'With frmGriege_Stock
            '        txtPer_Day.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            '        txtPer_Day.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            '        'End With
            '        '  Exit Function

            '    End If

            '    vcWhere = "left(M22Quality,2)in ('Y1') and M22Fabric_Type='SINGLE JERSEY' and M22Machine_Type not like '%Auto Stripe%' and M22Strich_Lenth<>0 and M22Quality='" & Trim(txtQuality.Text) & "' "
            '    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPER_DAY_KNTOUTPUT", New SqlParameter("@cQryType", "RPE2"), New SqlParameter("@vcWhereClause1", vcWhere))
            '    If isValidDataset(M01) Then
            '        Value = M01.Tables(0).Rows(0)("kgH")
            '        strPer_Day = Value
            '        'With frmGriege_Stock
            '        txtPer_Day.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            '        txtPer_Day.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            '        'End With
            '        ' Exit Function

            '    End If


            '    vcWhere = "left(M22Quality,1) not in ('Y') and M22Yarn_Cons='100' and M22Fabric_Type='SINGLE JERSEY' and M22Strich_Lenth<>0 and M22Quality='" & Trim(txtQuality.Text) & "' "
            '    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPER_DAY_KNTOUTPUT", New SqlParameter("@cQryType", "RPE2"), New SqlParameter("@vcWhereClause1", vcWhere))
            '    If isValidDataset(M01) Then
            '        Value = M01.Tables(0).Rows(0)("kgH")
            '        strPer_Day = Value
            '        ' With frmGriege_Stock
            '        txtPer_Day.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            '        txtPer_Day.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            '        'End With
            '        '  Exit Function

            '    End If

            '    vcWhere = "left(M22Quality,2)  in ('Y3') and M22Yarn_Cons<'100' and M22Fabric_Type='SINGLE JERSEY' and M22Machine_Type not like '%Auto Stripe%' and M22Strich_Lenth<>0 and M22Quality='" & Trim(txtQuality.Text) & "' "
            '    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPER_DAY_KNTOUTPUT", New SqlParameter("@cQryType", "RPE3"), New SqlParameter("@vcWhereClause1", vcWhere))
            '    If isValidDataset(M01) Then
            '        Value = M01.Tables(0).Rows(0)("kgH")
            '        strPer_Day = Value
            '        'With frmGriege_Stock
            '        txtPer_Day.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            '        txtPer_Day.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            '        'End With
            '        '   Exit Function

            '    End If

            '    vcWhere = "left(M22Quality,1) NOT in ('Y') and M22Yarn_Cons<'100' and M22Fabric_Type='SINGLE JERSEY' and M22Strich_Lenth<>0 and M22Quality='" & Trim(txtQuality.Text) & "' "
            '    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPER_DAY_KNTOUTPUT", New SqlParameter("@cQryType", "RPE3"), New SqlParameter("@vcWhereClause1", vcWhere))
            '    If isValidDataset(M01) Then
            '        Value = M01.Tables(0).Rows(0)("kgH")
            '        strPer_Day = Value
            '        'With frmGriege_Stock
            '        txtPer_Day.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            '        txtPer_Day.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            '        'End With
            '        ' Exit Function

            '    End If

            '    'End If

            'End If
            ''RIB

            'Dim _EFF As Double
            '_EFF = 0
            '' Value = txtPer_Day.Text
            'vcWhere = "left(M22Quality,2)in ('Y1','Y3') and M22Fabric_Type='SINGLE JERSEY' and M22Quality='" & Trim(txtQuality.Text) & "' "
            'M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPER_DAY_KNTOUTPUT", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWhere))
            'If isValidDataset(M01) Then
            '    _EFF = 0.6
            'End If

            'vcWhere = "M22Product_Type like '%MARL%' and left(M22Quality,2) not in ('Y1','Y3') and M22Fabric_Type='SINGLE JERSEY' and M22Quality='" & Trim(txtQuality.Text) & "' "
            'M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPER_DAY_KNTOUTPUT", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWhere))
            'If isValidDataset(M01) Then
            '    _EFF = 0.65
            'End If

            'vcWhere = "M22Product_Type not like '%MARL%' and left(M22Quality,2) not in ('Y1','Y3') and M22Fabric_Type='SINGLE JERSEY' and M22Quality='" & Trim(txtQuality.Text) & "' "
            'M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPER_DAY_KNTOUTPUT", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWhere))
            'If isValidDataset(M01) Then
            '    _EFF = 0.7
            'End If
            'If Value > 0 And _EFF > 0 Then
            '    Value = Value * _EFF * 24
            '    strPer_Day = Value
            '    'With frmGriege_Stock
            '    txtPer_Day.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            '    txtPer_Day.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            '    'End With
            'End If

            vcWhere = "M22Quality='" & txtQuality.Text & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TEC1"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                Value = M01.Tables(0).Rows(0)("M22Kg_Hr") * 24
            End If

            vcWhere = "M22Quality='" & txtQuality.Text & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "DDY"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                Value = Value * M01.Tables(0).Rows(0)("Knt_Eff")
                strPer_Day = Value
                txtPer_Day.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtPer_Day.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                con.close()
                Exit Function
            End If


            vcWhere = "M22Quality='" & txtQuality.Text & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "DJM"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                Value = Value * M01.Tables(0).Rows(0)("Knt_Eff")
                strPer_Day = Value
                txtPer_Day.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtPer_Day.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                con.close()
                Exit Function
            End If


            vcWhere = "M22Quality='" & txtQuality.Text & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "DJSL"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                Value = Value * M01.Tables(0).Rows(0)("Knt_Eff")
                strPer_Day = Value
                txtPer_Day.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtPer_Day.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                con.close()
                Exit Function
            End If


            vcWhere = "M22Quality='" & txtQuality.Text & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "SJS"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                Value = Value * M01.Tables(0).Rows(0)("Knt_Eff")
                strPer_Day = Value
                txtPer_Day.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtPer_Day.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                con.close()
                Exit Function
            End If


            vcWhere = "M22Quality='" & txtQuality.Text & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "SJA"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                Value = Value * M01.Tables(0).Rows(0)("Knt_Eff")
                strPer_Day = Value
                txtPer_Day.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtPer_Day.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                con.close()
                Exit Function
            End If


            vcWhere = "M22Quality='" & txtQuality.Text & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "SJDY"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                Value = Value * M01.Tables(0).Rows(0)("Knt_Eff")
                strPer_Day = Value
                txtPer_Day.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtPer_Day.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                con.close()
                Exit Function
            End If

            vcWhere = "M22Quality='" & txtQuality.Text & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "SJML"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                Value = Value * M01.Tables(0).Rows(0)("Knt_Eff")
                strPer_Day = Value
                txtPer_Day.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtPer_Day.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                con.close()
                Exit Function
            End If

            vcWhere = "M22Quality='" & txtQuality.Text & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "SJSO"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                Value = Value * M01.Tables(0).Rows(0)("Knt_Eff")
                strPer_Day = Value
                txtPer_Day.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtPer_Day.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                con.close()
                Exit Function
            End If

            vcWhere = "M22Quality='" & txtQuality.Text & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "DJSY"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                Value = Value * M01.Tables(0).Rows(0)("Knt_Eff")
                strPer_Day = Value
                txtPer_Day.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtPer_Day.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                con.close()
                Exit Function
            End If

            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                'Con.close()
            End If
        End Try
    End Function

    Private Sub frmKnitting_Plan_WithTab_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtFabric_Shade.ReadOnly = True
        Call Load_Gride()
        Call Load_GrideStock()
        txtDy_Year.Text = Year(Today)
        txtYear_Knt.Text = Year(Today)

        txtDy_Capacity.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
        txtDy_Week.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtDy_Year.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtAllocated_Qty.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtY_Balance.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

        txtY_Year.Text = Year(Today)
        txtY_Qty.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
        txtY_Week.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtY_Year.Appearance.TextHAlign = Infragistics.Win.HAlign.Center


        txtGriege_Qty.ReadOnly = True
        txtDelivary_Date.Text = Today
        txtFabric_Type.ReadOnly = True
        txtBasic_Yarn.ReadOnly = True
        txtNo_Colour.ReadOnly = True
        txtSpe_Yarn.ReadOnly = True
        ' txtShade.ReadOnly = True
        txtReq_Grg.ReadOnly = True
        txtDate.Text = Today

        Dim ToolTip1 As New ToolTip()
        ToolTip1.AutomaticDelay = 5000
        ToolTip1.InitialDelay = 1000
        ToolTip1.ReshowDelay = 500
        ToolTip1.ShowAlways = True
        Dim strTT As String
        ToolTip1.SetToolTip(cmdChart, cmdChart.Text & ControlChars.NewLine & "Graphical Yarn Plan")

        Dim ToolTip2 As New ToolTip()
        ToolTip2.AutomaticDelay = 5000
        ToolTip2.InitialDelay = 1000
        ToolTip2.ReshowDelay = 500
        ToolTip2.ShowAlways = True
        Dim strTT1 As String
        ToolTip2.SetToolTip(cmdDye_Yarn, cmdDye_Yarn.Text & ControlChars.NewLine & "Dyed Yarn Plan")

        Dim ToolTip3 As New ToolTip()
        ToolTip3.AutomaticDelay = 5000
        ToolTip3.InitialDelay = 1000
        ToolTip3.ReshowDelay = 500
        ToolTip3.ShowAlways = True
        Dim strTT3 As String
        ToolTip3.SetToolTip(cmdYarn_Request, cmdYarn_Request.Text & ControlChars.NewLine & "Yarn Request")

        Dim ToolTip4 As New ToolTip()
        ToolTip4.AutomaticDelay = 5000
        ToolTip4.InitialDelay = 1000
        ToolTip4.ReshowDelay = 500
        ToolTip4.ShowAlways = True
        Dim strTT4 As String
        ToolTip4.SetToolTip(cmdWinding, cmdWinding.Text & ControlChars.NewLine & "Soft Winding Plan")

        txt15Class.ReadOnly = True
        txtQty.ReadOnly = True

        txtQuality.ReadOnly = True
        txtQuality.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

        txtReq_Grg.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtK_Qty.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        'Call Load_Detailes()
        Call Load_GrideYD()
        ' Call Load_Gridewith_Data()
        ' Call Load_DataGD()
        Call Load_Gride_StockCode()
        Call Load_Gride_YarnStock()
        '  Call Delete_Transaction()
        Call Load_GrideDye_Plan()

        txtSales_Order.ReadOnly = True
        txtSales_Order.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtMC_Group_Knt.ReadOnly = True
        txtLine_Knt.ReadOnly = True
        txtLine_Knt.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtQuality_knt.ReadOnly = True
        txtQuality_knt.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtMC_Group_Knt.ReadOnly = True
        txtQty_Knt.ReadOnly = True
        txtQty_Knt.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtDaily_Capacity.ReadOnly = True
        txtDaily_Capacity.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtDays.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtDays.ReadOnly = True
        txtDelivary_Date.Text = Today
        txtComplete_Date_Knt.Text = Today

        Call Load_Gride_Knt()
        Call Load_Gride_projection()

        txtMonth1.ReadOnly = True
        txtMonth2.ReadOnly = True
        txtMonth3.ReadOnly = True

        txtMonth1.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtMonth2.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtMonth3.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

        txtPos1.ReadOnly = True
        txtPos2.ReadOnly = True
        txtPos3.ReadOnly = True

        txtPos1.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtPos2.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtPos3.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

        txtCon1.ReadOnly = True
        txtCon2.ReadOnly = True
        txtCon3.ReadOnly = True

        txtCon1.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtCon2.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtCon3.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtBal1.ReadOnly = True
        txtBal2.ReadOnly = True
        txtBal3.ReadOnly = True
        txtBal1.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtBal2.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtBal3.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

        txtPro1.ReadOnly = True
        txtPro2.ReadOnly = True
        txtPro3.ReadOnly = True

        txtPro1.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtPro2.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtPro3.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

        txtSL_Date.Text = Today
        txtSL_Line.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtSL_Qty.Appearance.TextHAlign = Infragistics.Win.HAlign.Center


        Call Load_Projection()

        Call Delete_Transaction()

        Call Load_Grid_Delivary_Plan()
    End Sub

    Function Load_Projection()
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim vcWhere As String
        Dim M01 As DataSet
        Dim _ProjectCode As Integer
        Dim Value As Double
        Dim _DDATE As Date
        Dim M15Project As DataSet
        Dim T01 As DataSet

        Try
            If Microsoft.VisualBasic.Day(Today) >= 10 Then
                _DDATE = Today.AddDays(+30)
                _DDATE = Month(_DDATE) & "/1/" & Year(_DDATE)
            Else
                _DDATE = Month(Today) & "/1/" & Year(Today)
            End If
            vcWhere = "P01CODE='PRN'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "P01"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                _ProjectCode = M01.Tables(0).Rows(0)("P01NO")
            End If
            If _ProjectCode > 0 Then
                _ProjectCode = _ProjectCode - 1
            End If
            i = 0
            Dim _StelingQulity As String
            'DEVELOPED BY SURANGA WIJESINGHE
            _StelingQulity = ""
            vcWhere = "T15Sales_Order='" & strSales_Order & "' AND T15Line_Item=" & strLine_Item & ""
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "BP1"), New SqlParameter("@vcWhereClause1", vcWhere))
            For Each DTRow As DataRow In M01.Tables(0).Rows
                If i = 0 Then
                    _StelingQulity = Trim(M01.Tables(0).Rows(i)("T15Quality"))
                Else
                    _StelingQulity = "','" & Trim(M01.Tables(0).Rows(i)("T15Quality"))
                End If
                i = i + 1
            Next
            i = 0
            ' MsgBox(Delivary_Ref)
            vcWhere = "M43Quality IN ('" & _StelingQulity & "') and M43Count_No=" & _ProjectCode & " "
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "PRS2"), New SqlParameter("@vcWhereClause1", vcWhere))

            For Each DTRow As DataRow In M01.Tables(0).Rows
                If i = 0 Then
                    txtMonth1.Text = MonthName(M01.Tables(0).Rows(i)("M43Product_Month"))
                    vcWhere = "M43Quality IN ('" & _StelingQulity & "') and M43Count_No=" & _ProjectCode & " and M43Year=" & M01.Tables(0).Rows(i)("M43Year") & " and M43Product_Month=" & M01.Tables(0).Rows(i)("M43Product_MontH") & ""
                    M15Project = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TMP5"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M15Project) Then
                        Value = M15Project.Tables(0).Rows(0)("QTY")
                        txtPro1.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        txtPro1.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    End If

                    Value = 0

                    vcWhere = "tmpQuality in ('" & _StelingQulity & "') and tmpYear=" & M01.Tables(0).Rows(i)("M43Year") & " and tmpMonth=" & M01.Tables(0).Rows(i)("M43Product_MontH") & ""
                    T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TMP1"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(T01) Then
                        Value = T01.Tables(0).Rows(0)("tmpQty")
                    End If
                    '===========================================================END STATMENT
                    'CHECK AVAILABLE ALLOCATED QTY FOR PROJECTION T15Projection

                    vcWhere = "T15Quality in ('" & _StelingQulity & "') and T15Year=" & M01.Tables(0).Rows(i)("M43Year") & " and T15Month=" & M01.Tables(0).Rows(i)("M43Product_Month") & " and T15Status='N' and T15Sales_Order<>'" & strSales_Order & "' "
                    T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TMP2"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(T01) Then
                        Value = Value + T01.Tables(0).Rows(0)("T15Qty")
                    End If

                    If Value > 0 Then
                        txtCon1.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        txtCon1.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    End If
                    Value = 0
                    If IsNumeric(txtPro1.Text) Then
                        Value = txtPro1.Text
                        pbCount.Maximum = Value

                    End If
                    If IsNumeric(txtCon1.Text) Then
                        pbCount.Value = txtCon1.Text
                    End If
                    If IsNumeric(txtCon1.Text) Then
                        Value = Value - txtCon1.Text
                        txtBal1.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        txtBal1.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    End If

                    If txtCon1.Text <> "" Then
                    Else
                        txtBal1.Text = txtPro1.Text
                        'txtBal1.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        'txtBal1.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    End If

                    vcWhere = "T15Sales_Order='" & strSales_Order & "' and T15Line_Item='" & strLine_Item & "' and T15Year=" & M01.Tables(0).Rows(i)("M43Year") & " and T15Month=" & M01.Tables(0).Rows(i)("M43Product_MontH") & ""
                    M15Project = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TMP6"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M15Project) Then
                        Value = M15Project.Tables(0).Rows(0)("Qty")
                        txtPos1.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        txtPos1.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    End If

                ElseIf i = 1 Then
                    txtMonth2.Text = MonthName(M01.Tables(0).Rows(i)("M43Product_Month"))
                    vcWhere = "M43Quality IN ('" & _StelingQulity & "') and M43Count_No=" & _ProjectCode & " and M43Year=" & M01.Tables(0).Rows(i)("M43Year") & " and M43Product_Month=" & M01.Tables(0).Rows(i)("M43Product_MontH") & ""
                    M15Project = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TMP5"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M15Project) Then
                        Value = M15Project.Tables(0).Rows(0)("QTY")
                        txtPro2.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        txtPro2.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    End If

                    Value = 0

                    vcWhere = "tmpQuality in ('" & _StelingQulity & "') and tmpYear=" & M01.Tables(0).Rows(i)("M43Year") & " and tmpMonth=" & M01.Tables(0).Rows(i)("M43Product_MontH") & ""
                    T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TMP1"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(T01) Then
                        Value = T01.Tables(0).Rows(0)("tmpQty")
                    End If
                    '===========================================================END STATMENT
                    'CHECK AVAILABLE ALLOCATED QTY FOR PROJECTION T15Projection

                    vcWhere = "T15Quality in ('" & _StelingQulity & "') and T15Year=" & M01.Tables(0).Rows(i)("M43Year") & " and T15Month=" & M01.Tables(0).Rows(i)("M43Product_Month") & " and T15Status='N' and T15Sales_Order<>'" & strSales_Order & "'"
                    T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TMP2"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(T01) Then
                        Value = Value + T01.Tables(0).Rows(0)("T15Qty")
                    End If

                    If Value > 0 Then
                        txtCon2.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        txtCon2.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    End If

                    Value = 0
                    If IsNumeric(txtPro2.Text) Then
                        Value = txtPro2.Text
                        pbcount1.Maximum = Value
                    End If
                    If IsNumeric(txtCon2.Text) Then
                        Value = Value - txtCon2.Text
                        txtBal2.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        txtBal2.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                        pbcount1.Value = txtCon2.Text
                    End If

                    If txtCon2.Text <> "" Then
                    Else
                        txtBal2.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        txtBal2.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    End If
                    If txtCon2.Text <> "" Then
                    Else
                        pbcount1.Value = txtPro2.Value
                    End If

                    vcWhere = "T15Sales_Order='" & strSales_Order & "' and T15Line_Item='" & strLine_Item & "' and T15Year=" & M01.Tables(0).Rows(i)("M43Year") & " and T15Month=" & M01.Tables(0).Rows(i)("M43Product_MontH") & ""
                    M15Project = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TMP6"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M15Project) Then
                        Value = M15Project.Tables(0).Rows(0)("Qty")
                        txtPos2.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        txtPos2.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    End If

                ElseIf i = 2 Then
                    txtMonth3.Text = MonthName(M01.Tables(0).Rows(i)("M43Product_Month"))
                    vcWhere = "M43Quality IN ('" & _StelingQulity & "') and M43Count_No=" & _ProjectCode & " and M43Year=" & M01.Tables(0).Rows(i)("M43Year") & " and M43Product_Month=" & M01.Tables(0).Rows(i)("M43Product_MontH") & ""
                    M15Project = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TMP5"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M15Project) Then
                        Value = M15Project.Tables(0).Rows(0)("QTY")
                        txtPro3.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        txtPro3.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    End If

                    Value = 0

                    vcWhere = "tmpQuality in ('" & _StelingQulity & "') and tmpYear=" & M01.Tables(0).Rows(i)("M43Year") & " and tmpMonth=" & M01.Tables(0).Rows(i)("M43Product_MontH") & ""
                    T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TMP1"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(T01) Then
                        Value = T01.Tables(0).Rows(0)("tmpQty")
                    End If
                    '===========================================================END STATMENT
                    'CHECK AVAILABLE ALLOCATED QTY FOR PROJECTION T15Projection

                    vcWhere = "T15Quality in ('" & _StelingQulity & "') and T15Year=" & M01.Tables(0).Rows(i)("M43Year") & " and T15Month=" & M01.Tables(0).Rows(i)("M43Product_Month") & " and T15Status='N' and T15Sales_Order<>'" & strSales_Order & "'"
                    T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TMP2"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(T01) Then
                        Value = Value + T01.Tables(0).Rows(0)("T15Qty")
                    End If

                    If Value > 0 Then
                        txtCon3.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        txtCon3.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    End If


                    Value = 0
                    If IsNumeric(txtPro3.Text) Then
                        Value = txtPro3.Text
                        pbCount2.Maximum = Value
                    End If
                    If IsNumeric(txtCon3.Text) Then
                        Value = Value - txtCon3.Text
                        txtBal3.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        txtBal3.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                        pbCount2.Value = txtCon3.Text
                    End If


                    If txtCon3.Text <> "" Then

                        txtBal3.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        txtBal3.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    Else
                        txtBal3.Text = txtPro3.Text
                        pbCount2.Value = txtPro3.Text
                    End If

                    vcWhere = "T15Sales_Order='" & strSales_Order & "' and T15Line_Item='" & strLine_Item & "' and T15Year=" & M01.Tables(0).Rows(i)("M43Year") & " and T15Month=" & M01.Tables(0).Rows(i)("M43Product_MontH") & ""
                    M15Project = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TMP6"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M15Project) Then
                        Value = M15Project.Tables(0).Rows(0)("Qty")
                        txtPos3.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        txtPos3.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    End If

                End If
                i = i + 1
            Next

            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            con.close()
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)

                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
                con.close()
            End If
        End Try
    End Function

    Function Load_Gride_StockCode()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer3 = CustomerDataClass.MakeDataTableYarn_Stock
        dg2.DataSource = c_dataCustomer3
        With dg2
            .DisplayLayout.Bands(0).Columns(0).Width = 40
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = True
            .DisplayLayout.Bands(0).Columns(1).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = True
            .DisplayLayout.Bands(0).Columns(2).Width = 160
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = True
            .DisplayLayout.Bands(0).Columns(3).Width = 70
            .DisplayLayout.Bands(0).Columns(3).AutoEdit = True
            .DisplayLayout.Bands(0).Columns(4).Width = 70
            .DisplayLayout.Bands(0).Columns(5).Width = 70

            .DisplayLayout.Bands(0).Columns(5).AutoEdit = False
            '   .DisplayLayout.Bands(0).Columns(6).AutoEdit = False
            ' .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Left
            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            '.DisplayLayout.Bands(0).Columns(3).Width = 90
            '.DisplayLayout.Bands(0).Columns(4).Width = 90
            '.DisplayLayout.Bands(0).Columns(5).Width = 90
            ''  .DisplayLayout.Bands(0).Columns(6).Width = 90
            ' .DisplayLayout.Bands(0).Columns(7).Width = 90

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function

    Function Load_Gridewith_Data()
        Dim i As Integer
        Dim vcWhere As String
        Dim M01 As DataSet
        Dim M02 As DataSet
        Dim Value As Double
        Dim _VString As String
        Dim Diff As TimeSpan
        Dim _To As Date
        'Dim Value As Double
        Dim _Rowcount As Integer
        Dim characterToRemove As String

        Try
            Dim con = New SqlConnection()
            con = DBEngin.GetConnection(True)


            Dim Z As Integer
            Z = 0
            i = 0
            vcWhere = "M22Quality='" & Trim(txtQuality.Text) & "' and left(M22M_Class,2)='15'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TEC1"), New SqlParameter("@vcWhereClause1", vcWhere))
            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer2.NewRow

                Dim _STValue As String
                Dim _Rcode As String

                newRow("15Class") = M01.Tables(0).Rows(i)("M22M_Class")
                newRow("Description") = M01.Tables(0).Rows(i)("M22Yarn")
                newRow("Composition") = CDbl(M01.Tables(0).Rows(i)("M22Yarn_Cons"))
                _Rcode = ""
                _Rcode = Microsoft.VisualBasic.Right(M01.Tables(0).Rows(i)("M22Yarn"), 6)
                _Rcode = Microsoft.VisualBasic.Left(_Rcode, 5)

                characterToRemove = "Y"
                _Rcode = (Replace(_Rcode, characterToRemove, ""))
                _Rcode = Trim(_Rcode)
                vcWhere = "M14Order='" & _Rcode & "' and M14Type='Y'"
                M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "RCDE1"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M02) Then
                    If Trim(M02.Tables(0).Rows(0)("M14Shade_Cat")) = "D" Then
                        newRow("Shade") = "DARK"
                    ElseIf Trim(M02.Tables(0).Rows(0)("M14Shade_Cat")) = "L" Then
                        newRow("Shade") = "LIGHT"
                    ElseIf Trim(M02.Tables(0).Rows(0)("M14Shade_Cat")) = "M" Then
                        newRow("Shade") = "MARL"
                    ElseIf Trim(M02.Tables(0).Rows(0)("M14Shade_Cat")) = "MARL" Then
                        newRow("Shade") = "MARL"
                    End If
                End If
                Value = 0
                If IsNumeric(txtReq_Grg.Text) Then
                    Value = CDbl(txtReq_Grg.Text)
                    If IsNumeric(txtDye_Wast.Text) Then
                        Value = Value / ((100 - txtDye_Wast.Text) / 100)
                    End If
                    Value = Value * CDbl(M01.Tables(0).Rows(i)("M22Yarn_Cons"))
                    Value = Value / 100

                    _STValue = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    _STValue = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                End If
                newRow("Dyed Yarn Req.Knit") = _STValue
                _STValue = ""
                If IsNumeric(txtYarn_Wst.Text) Then
                    Value = Value / ((100 - txtYarn_Wst.Text) / 100)
                    _STValue = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    _STValue = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                End If
                newRow("Yarn Req - Dyed Yarn") = _STValue
                newRow("Balance Qty for Allocate") = _STValue
                newRow("No of Cons") = CInt(Value / 1.05)
                c_dataCustomer2.Rows.Add(newRow)


                i = i + 1
            Next
            Dim newRow1 As DataRow = c_dataCustomer2.NewRow
            c_dataCustomer2.Rows.Add(newRow1)
            _Rowcount = dg1.Rows.Count
            dg1.Rows(_Rowcount - 1).Cells(0).Appearance.BackColor = Color.DeepSkyBlue
            dg1.Rows(_Rowcount - 1).Cells(1).Appearance.BackColor = Color.DeepSkyBlue
            dg1.Rows(_Rowcount - 1).Cells(2).Appearance.BackColor = Color.DeepSkyBlue
            dg1.Rows(_Rowcount - 1).Cells(3).Appearance.BackColor = Color.DeepSkyBlue
            dg1.Rows(_Rowcount - 1).Cells(4).Appearance.BackColor = Color.DeepSkyBlue
            dg1.Rows(_Rowcount - 1).Cells(5).Appearance.BackColor = Color.DeepSkyBlue
            dg1.Rows(_Rowcount - 1).Cells(6).Appearance.BackColor = Color.DeepSkyBlue
            dg1.Rows(_Rowcount - 1).Cells(7).Appearance.BackColor = Color.DeepSkyBlue
            dg1.Rows(_Rowcount - 1).Cells(8).Appearance.BackColor = Color.DeepSkyBlue
            dg1.Rows(_Rowcount - 1).Cells(9).Appearance.BackColor = Color.DeepSkyBlue

            i = 0
            vcWhere = "M22Quality='" & Trim(txtQuality.Text) & "' and left(M22M_Class,2)<>'15'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TEC"), New SqlParameter("@vcWhereClause1", vcWhere))
            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer2.NewRow
                Dim _STValue As String

                _STValue = ""
                ' newRow("15Class") = M01.Tables(0).Rows(i)("M22M_Class")
                newRow("Description") = M01.Tables(0).Rows(i)("M22Yarn")
                newRow("Composition") = CDbl(M01.Tables(0).Rows(i)("M22Yarn_Cons"))
                If IsNumeric(txtReq_Grg.Text) Then
                    Value = CDbl(txtReq_Grg.Text)
                    If IsNumeric(txtDY_To_Greige.Text) Then
                        Value = Value / 0.98
                    End If
                    Value = Value * CDbl(M01.Tables(0).Rows(i)("M22Yarn_Cons"))
                    Value = Value / 100

                    _STValue = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    _STValue = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                End If
                newRow("Dyed Yarn Req.Knit") = _STValue
                newRow("Yarn Req - Dyed Yarn") = _STValue
                c_dataCustomer2.Rows.Add(newRow)


                i = i + 1
            Next

            Dim newRow2 As DataRow = c_dataCustomer2.NewRow
            c_dataCustomer2.Rows.Add(newRow2)
            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                'Con.close()
            End If
        End Try
    End Function

    Function Load_Detailes()
        Dim i As Integer
        Dim vcWhere As String
        Dim M01 As DataSet
        Dim M02 As DataSet
        Dim Value As Double
        Dim _VString As String
        Dim Diff As TimeSpan
        Dim _To As Date
        Dim MyText As String

        Try
            Dim con = New SqlConnection()
            con = DBEngin.GetConnection(True)

            txtReq_Greige.Text = txtGriege_Qty.Text
            txtQuality_YD.Text = txtQuality.Text
            i = 0
            vcWhere = "M22Quality='" & Trim(txtQuality.Text) & "' and left(M22M_Class,2)='15'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TEC"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                txtNo_Colour.Text = M01.Tables(0).Rows.Count
                MyText = M01.Tables(0).Rows(0)("M22Yarn")
                Dim myIndex = MyText.IndexOf("")
                txtBasic_Yarn.Text = Microsoft.VisualBasic.Left(M01.Tables(0).Rows(0)("M22Yarn"), myIndex)
            End If

            vcWhere = "M22Quality='" & Trim(txtQuality.Text) & "' and left(M22M_Class,2)<>'15'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TEC"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                If M01.Tables(0).Rows(0)("M22Strich_Lenth") < 0 Then
                    chkNPL1.Checked = True
                    txtSpe_Yarn.Text = M01.Tables(0).Rows(0)("M22Yarn")

                Else
                    chkNPL2.Checked = True
                End If
            End If

            txtWastage.Text = txtDye_Wast.Text
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetM28Stock_Grige_Price", New SqlParameter("@cQryType", "DWS"))
            If isValidDataset(M01) Then
                txtDY_To_Greige.Text = M01.Tables(0).Rows(0)("M35D_WST")
                txtYarn_Wst.Text = M01.Tables(0).Rows(0)("M35Y_WST")

            End If

            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                'Con.close()
            End If
        End Try
    End Function

    Function Load_GrideYD()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer2 = CustomerDataClass.MakeDataTableYarn_Dyeing
        dg1.DataSource = c_dataCustomer2
        With dg1
            .DisplayLayout.Bands(0).Columns(0).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 230
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 80
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(3).Width = 90
            .DisplayLayout.Bands(0).Columns(3).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(4).Width = 90
            .DisplayLayout.Bands(0).Columns(4).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(5).Width = 110
            .DisplayLayout.Bands(0).Columns(5).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(5).Style = ColumnStyle.EditButton
            .DisplayLayout.Bands(0).Columns(6).Width = 110
            .DisplayLayout.Bands(0).Columns(6).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(9).Style = ColumnStyle.EditButton
            '   .DisplayLayout.Bands(0).Columns(6).AutoEdit = False
            ' .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Left
            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Left
            .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(9).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(8).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

            '.DisplayLayout.Bands(0).Columns(3).Width = 90
            '.DisplayLayout.Bands(0).Columns(4).Width = 90
            '.DisplayLayout.Bands(0).Columns(5).Width = 90
            ''  .DisplayLayout.Bands(0).Columns(6).Width = 90
            ' .DisplayLayout.Bands(0).Columns(7).Width = 90

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function

    Function Load_Gride_YarnStock()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer4 = CustomerDataClass.MakeDataTableYarn
        dg4.DataSource = c_dataCustomer4
        With dg4
            .DisplayLayout.Bands(0).Columns(0).Width = 40
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = True
            .DisplayLayout.Bands(0).Columns(1).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = True
            .DisplayLayout.Bands(0).Columns(2).Width = 250
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = True
            .DisplayLayout.Bands(0).Columns(3).Width = 70
            .DisplayLayout.Bands(0).Columns(3).AutoEdit = True
            .DisplayLayout.Bands(0).Columns(4).Width = 70
            .DisplayLayout.Bands(0).Columns(5).Width = 70
            .DisplayLayout.Bands(0).Columns(5).AutoEdit = False
            '   .DisplayLayout.Bands(0).Columns(6).AutoEdit = False
            ' .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Left
            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Left
            .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

            '.DisplayLayout.Bands(0).Columns(3).Width = 90
            '.DisplayLayout.Bands(0).Columns(4).Width = 90
            '.DisplayLayout.Bands(0).Columns(5).Width = 90
            ''  .DisplayLayout.Bands(0).Columns(6).Width = 90
            ' .DisplayLayout.Bands(0).Columns(7).Width = 90

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function

    Function Delete_Transaction_Yarn_Booking()
        Dim vcWhere As String
        Dim M01 As DataSet
        Dim vcFieldList As String
        Dim ncQryType As String
        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Try

            nvcFieldList1 = "delete from tmpBlock_Yarn_Stock_Code where tmpUser='" & strDisname & "'"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            transaction.Commit()
            connection.Close()
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)

                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
                connection.Close()
            End If
        End Try
    End Function

    Function Load_Grid_SockCode()
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim vcWhere As String
        Dim M01 As DataSet
        Dim i As Integer
        Dim Value As Double
        Dim _VString As String

        Try
            i = 0
            vcWhere = "M42Rcode='" & Trim(txtRcode.Text) & "'  "
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetUse_Griege_Qty", New SqlParameter("@cQryType", "STC"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
            Else
                vcWhere = "m42Quality='" & Trim(txtCommon.Text) & "'  "
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetUse_Griege_Qty", New SqlParameter("@cQryType", "STC"), New SqlParameter("@vcWhereClause1", vcWhere))
            End If
            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomerSTC.NewRow

                newRow("Stock Code") = M01.Tables(0).Rows(i)("M42Stock_Code")
                newRow("Date") = Month(M01.Tables(0).Rows(i)("M24Customer")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("M24Customer")) & "/" & Year(M01.Tables(0).Rows(i)("M24Customer"))
                newRow("Week No") = M01.Tables(0).Rows(i)("M24Week")
                newRow("Year") = M01.Tables(0).Rows(i)("M24Year")
                Value = M01.Tables(0).Rows(i)("Qty")
                _VString = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _VString = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Used Qty(Kg)") = _VString
                c_dataCustomerSTC.Rows.Add(newRow)
                i = i + 1
            Next
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
                con.close()
            End If
        End Try

    End Function

    Function Load_WithData()
        Dim i As Integer
        Dim vcWhere As String
        Dim M01 As DataSet
        Dim M02 As DataSet
        Dim Value As Double
        Dim _VString As String
        Dim Diff As TimeSpan
        Dim _To As Date
        Dim X As Integer
        Dim _Date As Date
        Dim _BatchNo As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim _qty1 As Double

        Try
          
            i = 0

            If Microsoft.VisualBasic.Left(Trim(txtFabric_Type.Text), 1) = "M" Or Microsoft.VisualBasic.Left(Trim(txtFabric_Type.Text), 1) = "P" Or Microsoft.VisualBasic.Left(Trim(txtFabric_Type.Text), 1) = "Y" Then
                vcWhere = "M21Material='" & Trim(txtQuality.Text) & "'" '  and left(M21Sales_Order,2)='20' "
            Else
                vcWhere = "M21Material='" & Trim(txtQuality.Text) & "' and left(M23Shade,1)='" & Trim(txtFabric_Type.Text) & "'" ' and left(M21Sales_Order,2)='20' "
            End If
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetUse_Griege_Qty", New SqlParameter("@cQryType", "UGS1"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
            Else
                vcWhere = "M21Material='" & Trim(txtCommon.Text) & "' and left(M23Shade,1)='" & Trim(txtShade.Text) & "'" ' and left(M21Sales_Order,2)='20' "
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetUse_Griege_Qty", New SqlParameter("@cQryType", "UGS1"), New SqlParameter("@vcWhereClause1", vcWhere))
            End If
            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                _qty1 = 0
                vcWhere = "T12Stock_Code='" & M01.Tables(0).Rows(i)("M21Batch_No") & "' and T12Time>='" & M01.Tables(0).Rows(i)("M21Update_Time") & "'  "
                M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetUse_Griege_Qty", New SqlParameter("@cQryType", "CGS"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M02) Then
                    _qty1 = M02.Tables(0).Rows(0)("T12Qty")
                End If
                If M01.Tables(0).Rows(i)("M21Qty") > _qty1 Then
                    Dim newRow As DataRow = c_dataCustomer1.NewRow

                    newRow("20Class") = M01.Tables(0).Rows(i)("M2120Class")
                    newRow("L/Item") = M01.Tables(0).Rows(i)("M21Line_Item")
                    newRow("Grige Order No") = M01.Tables(0).Rows(i)("M21Sales_Order")
                    newRow("Stock Code") = M01.Tables(0).Rows(i)("M21Batch_No")
                    If i = 0 Then
                        _BatchNo = "" & M01.Tables(0).Rows(i)("M21Batch_No")
                    Else
                        _BatchNo = _BatchNo & "','" & M01.Tables(0).Rows(i)("M21Batch_No")
                    End If
                    _To = Month(M01.Tables(0).Rows(i)("M21Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("M21Date")) & "/" & Year(M01.Tables(0).Rows(i)("M21Date"))
                    Diff = Today.Subtract(_To)
                    newRow("Age") = Diff.Days & " days"

                    'If Diff.Days < 30 Then
                    '    newRow("Age") = "Below 1 Month"
                    'ElseIf Diff.Days >= 30 And Diff.Days < 60 Then
                    '    newRow("Age") = "Below 2 Month"
                    'ElseIf Diff.Days >= 60 And Diff.Days < 90 Then
                    '    newRow("Age") = "Below 3 Month"
                    'Else
                    '    newRow("Age") = "above 3 Month"
                    'End If
                    Value = CDbl(M01.Tables(0).Rows(i)("M21Qty")) - _qty1
                    _VString = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    _VString = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    newRow("Available Qty(Kg)") = _VString
                    newRow("##") = False
                    c_dataCustomer1.Rows.Add(newRow)
                End If

                i = i + 1
            Next
            i = 0
            vcWhere = "M21Material='" & Trim(txtQuality.Text) & "' and left(M23Shade,1)='" & Trim(txtShade.Text) & "' and left(M21Sales_Order,2)='20' and M21Batch_No not in ('" & _BatchNo & "') "
            vcWhere = "M21Material='" & Trim(txtQuality.Text) & "' and left(M23Shade,1)='" & Trim(txtShade.Text) & "' and M21Batch_No not in ('" & _BatchNo & "') "
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetUse_Griege_Qty", New SqlParameter("@cQryType", "UGS"), New SqlParameter("@vcWhereClause1", vcWhere))
            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                _qty1 = 0
                vcWhere = "T12Stock_Code='" & M01.Tables(0).Rows(i)("M21Batch_No") & "' and T12Time>='" & M01.Tables(0).Rows(i)("M21Update_Time") & "'  "
                M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetUse_Griege_Qty", New SqlParameter("@cQryType", "CGS"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M02) Then
                    _qty1 = M02.Tables(0).Rows(0)("T12Qty")
                End If
                If M01.Tables(0).Rows(i)("M21Qty") > _qty1 Then
                    Dim newRow As DataRow = c_dataCustomer1.NewRow

                    newRow("20Class") = M01.Tables(0).Rows(i)("M2120Class")
                    newRow("Grige Order No") = M01.Tables(0).Rows(i)("M21Sales_Order")
                    newRow("Stock Code") = M01.Tables(0).Rows(i)("M21Batch_No")
                    _To = Month(M01.Tables(0).Rows(i)("M21Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("M21Date")) & "/" & Year(M01.Tables(0).Rows(i)("M21Date"))
                    Diff = Today.Subtract(_To)
                    newRow("Age") = Diff.Days & " days"
                    newRow("L/Item") = M01.Tables(0).Rows(i)("M21Line_Item")
                    'If Diff.Days < 30 Then
                    '    newRow("Age") = "Below 1 Month"
                    'ElseIf Diff.Days >= 30 And Diff.Days < 60 Then
                    '    newRow("Age") = "Below 2 Month"
                    'ElseIf Diff.Days >= 60 And Diff.Days < 90 Then
                    '    newRow("Age") = "Below 3 Month"
                    'Else
                    '    newRow("Age") = "above 3 Month"
                    'End If
                    Value = CDbl(M01.Tables(0).Rows(i)("M21Qty")) - _qty1
                    _VString = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    _VString = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    newRow("Available Qty(Kg)") = _VString
                    newRow("##") = False
                    c_dataCustomer1.Rows.Add(newRow)
                End If
                i = i + 1
            Next

            'con.close()
            X = 0
            For Each uRow As UltraGridRow In UltraGrid2.Rows
                i = 0
                'If lblBalance.Text = "0.00" Then
                '    Exit Function
                'End If
                For Each uRow1 As UltraGridRow In UltraGrid1.Rows
                    Dim _Qty As Double
                    _Qty = 0
                    With UltraGrid1
                        If Trim(.Rows(i).Cells(3).Value) = Trim(UltraGrid2.Rows(X).Cells(0).Value) Then
                            _Qty = .Rows(i).Cells(5).Value
                            'If CDbl(lblBalance.Text) >= _Qty Then
                            'If UltraGrid1.Rows(i).Cells(8).Value = True Then
                            'Else
                            ' lblBalance.Text = CDbl(lblBalance.Text) - _Qty
                            .Rows(i).Cells(0).Appearance.BackColor = Color.Blue
                            .Rows(i).Cells(1).Appearance.BackColor = Color.Blue
                            .Rows(i).Cells(2).Appearance.BackColor = Color.Blue
                            .Rows(i).Cells(3).Appearance.BackColor = Color.Blue
                            .Rows(i).Cells(4).Appearance.BackColor = Color.Blue
                            .Rows(i).Cells(5).Appearance.BackColor = Color.Blue
                            .Rows(i).Cells(6).Appearance.BackColor = Color.Blue
                            .Rows(i).Cells(7).Appearance.BackColor = Color.Blue
                            .Rows(i).Cells(8).Appearance.BackColor = Color.Blue
                            ' .Rows(i).Cells(5).Text = lblBalance.Text
                            .Rows(i).Cells(6).Value = _Qty
                            '    .Rows(i).Cells(8).Value = True
                            'End If
                            'Else
                            'If UltraGrid1.Rows(i).Cells(8).Value = True Then
                            'Else
                            '    If CDbl(lblBalance.Text) = "0.00" Then
                            '    Else

                            '        .Rows(i).Cells(0).Appearance.BackColor = Color.Blue
                            '        .Rows(i).Cells(1).Appearance.BackColor = Color.Blue
                            '        .Rows(i).Cells(2).Appearance.BackColor = Color.Blue
                            '        .Rows(i).Cells(3).Appearance.BackColor = Color.Blue
                            '        .Rows(i).Cells(4).Appearance.BackColor = Color.Blue
                            '        .Rows(i).Cells(5).Appearance.BackColor = Color.Blue
                            '        .Rows(i).Cells(6).Appearance.BackColor = Color.Blue
                            '        .Rows(i).Cells(7).Appearance.BackColor = Color.Blue
                            '        .Rows(i).Cells(8).Appearance.BackColor = Color.Blue
                            '        ' .Rows(i).Cells(5).Text = lblBalance.Text
                            '        .Rows(i).Cells(5).Value = lblBalance.Text
                            '        .Rows(i).Cells(8).Value = True
                            '        lblBalance.Text = "0.00"
                            '        Exit Function
                            '    End If
                            'End If
                            'End If
                        End If


                    End With
                    i = i + 1
                Next
                X = X + 1
            Next
            ' con.close()
            i = 0
            For Each uRow1 As UltraGridRow In UltraGrid1.Rows
                If Microsoft.VisualBasic.Left(UltraGrid1.Rows(i).Cells(1).Value, 1) = "2" Then

                Else
                    vcWhere = "select * from OTD_Records where Sales_Order='" & UltraGrid1.Rows(i).Cells(1).Value & "' And Line_Item='" & UltraGrid1.Rows(i).Cells(2).Value & "' order by Del_Date desc"
                    M01 = DBEngin.ExecuteDataset(con, Nothing, vcWhere)
                    If isValidDataset(M01) Then
                        UltraGrid1.Rows(i).Cells(7).Value = Month(M01.Tables(0).Rows(0)("Del_Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(0)("Del_Date")) & "/" & Year(M01.Tables(0).Rows(0)("Del_Date"))
                    Else
                        UltraGrid1.Rows(i).Cells(7).Value = "Completed Order"
                    End If
                End If
                i = i + 1
            Next

            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                Con.close()
            End If
        End Try
    End Function

    Function Load_GrideDye_Plan()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer5 = CustomerDataClass.MakeDataTableDye_Plan
        dg3.DataSource = c_dataCustomer5
        With dg3
            .DisplayLayout.Bands(0).Columns(0).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 230
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 80
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(3).Width = 90
            .DisplayLayout.Bands(0).Columns(3).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(4).Width = 90
            .DisplayLayout.Bands(0).Columns(4).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(5).Width = 110

        End With
    End Function

    Function Load_Gride_projection()
        'Dim CustomerDataClass As New DAL_InterLocation()
        'c_dataCustomer2_KNT = CustomerDataClass.MakeDataTableKnt_Projection
        'dg1_Knt.DataSource = c_dataCustomer2_KNT
        'With dg1_Knt
        '    .DisplayLayout.Bands(0).Columns(0).Width = 170
        '    .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
        '    .DisplayLayout.Bands(0).Columns(1).Width = 70
        '    .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
        '    .DisplayLayout.Bands(0).Columns(2).Width = 70
        '    .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
        '    .DisplayLayout.Bands(0).Columns(3).Width = 70
        '    .DisplayLayout.Bands(0).Columns(3).AutoEdit = False
        '    .DisplayLayout.Bands(0).Columns(4).Width = 70
        '    .DisplayLayout.Bands(0).Columns(4).AutoEdit = False

        '    .DisplayLayout.Bands(0).Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
        '    .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
        '    .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
        '    .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
        '    .DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
        '    '.DisplayLayout.Bands(0).Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
        '    '.DisplayLayout.Bands(0).Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
        '    '.DisplayLayout.Bands(0).Columns(3).Width = 90
        '    '.DisplayLayout.Bands(0).Columns(4).Width = 90
        '    '.DisplayLayout.Bands(0).Columns(5).Width = 90
        '    ''  .DisplayLayout.Bands(0).Columns(6).Width = 90
        '    ' .DisplayLayout.Bands(0).Columns(7).Width = 90

        '    ' .DisplayLayout.Bands(0).Columns(3).Width = 300
        '    '.DisplayLayout.Bands(0).Columns(4).Width = 300
        'End With
    End Function

    Private Sub txtEx_LibUse_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        Call frmDelivaryQuatnew.CalculateBalance_To_Produce()
    End Sub

    Private Sub txtDate_TextChanged1(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDate.TextChanged
        Call Update_Date()
    End Sub

    Private Sub chkPP1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkPP1.CheckedChanged
        If chkPP1.Checked = True Then
            chkPP2.Checked = False
        End If
    End Sub

    Private Sub chkPP2_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkPP2.CheckedChanged
        If chkPP2.Checked = True Then
            chkPP1.Checked = False
        End If
    End Sub

    Private Sub chkCry1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkCry1.CheckedChanged
        If chkCry1.Checked = True Then
            chkCry2.Checked = False
        End If
    End Sub

  
    Private Sub chkCry2_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkCry2.CheckedChanged
        If chkCry2.Checked = True Then
            chkCry1.Checked = False
        End If
    End Sub

    Private Sub chkLab1_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkLab1.CheckedChanged
        If chkLab1.Checked = True Then
            chkLab2.Checked = False
        End If
    End Sub

    Private Sub chkLab2_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkLab2.CheckedChanged
        If chkLab2.Checked = True Then
            chkLab1.Checked = False
        End If
    End Sub

    Private Sub chkNPL1_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkNPL1.CheckedChanged
        If chkNPL1.Checked = True Then
            chkNPL2.Checked = False
        End If
    End Sub

    Private Sub chkNPL2_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkNPL2.CheckedChanged
        If chkNPL2.Checked = True Then
            chkNPL1.Checked = False
        End If
    End Sub

 
    Private Sub chkG_Stock_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkG_Stock.CheckedChanged
        Dim Value As Double
        Dim A As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim m01 As DataSet
        Dim VCWHERE As String
        Dim I As Integer
        Dim _Stcode As String

        If chkLab1.Checked = True Then
        Else
            VCWHERE = "T11Ref_No=" & Delivary_Ref & " and T11Sales_Order='" & txtSO.Text & "' and T11Line_Item=" & txtLine_Item.Text & ""
            m01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "LADC"), New SqlParameter("@vcWhereClause1", VCWHERE))
            If isValidDataset(m01) Then
            Else
                A = MsgBox("LD is not Approved are you want to  continue.", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Information ....")
                If A = vbYes Then
                    txtDate.Visible = True
                    txtDate.Text = Today
                    Exit Sub
                Else
                    Exit Sub
                End If
            End If
        End If

        '================================================================
        'Matching Stock Code
        _Stcode = ""
        VCWHERE = "T01Sales_Order='" & strSales_Order & "' and T01Line_Item=" & strLine_Item & ""
        m01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "T01D"), New SqlParameter("@vcWhereClause1", VCWHERE))
        If isValidDataset(m01) Then
            VCWHERE = "T01Sales_Order='" & strSales_Order & "' and t01Maching=" & strLine_Item & ""
            dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "T01D"), New SqlParameter("@vcWhereClause1", VCWHERE))
            If isValidDataset(dsUser) Then
                VCWHERE = "T12Sales_Order='" & strSales_Order & "' and T12Line_Item=" & Trim(dsUser.Tables(0).Rows(0)("T01Line_Item")) & ""
                dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStock_Grige", New SqlParameter("@cQryType", "GSC"), New SqlParameter("@vcWhereClause1", VCWHERE))
                I = 0
                For Each DTRow5 As DataRow In dsUser.Tables(0).Rows
                    If I = 0 Then
                        _Stcode = dsUser.Tables(0).Rows(I)("T12Stock_Code")
                    Else
                        _Stcode = _Stcode & " | " + dsUser.Tables(0).Rows(I)("T12Stock_Code")
                    End If
                    I = I + 1
                Next
            End If
        End If
        txtMC_Group.Text = _Stcode

        con.CLOSE()
        If chkG_Stock.Checked = True Then
            chkY_Orde.Checked = False
            chkY_Orde.Checked = False
            If txtReg_LIb.Text <> "" Then
            Else
                txtReg_LIb.Text = "0"
            End If
            If txtReq_Grg.Text <> "" Then
            Else
                txtReq_Grg.Text = "0"
            End If
            Value = CDbl(txtReq_Grg.Text) + CDbl(txtReg_LIb.Text)
            ' With frmGriege_Stock
            txtGriege_Qty.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            txtGriege_Qty.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            txtFabric.Text = txtFabrication.Text
            txtLIB.Text = txtLIB.Text
            'End With
            Call Search_Data()

            Call Load_Parameter()
            Call Update_Parameter()
            'If Trim(txtFabric_Shade.Text) = "Yarn Dyes" Then
            '    chkKnt_Plan.Text = "Yarn Dye Plan"
            'End If
            UltraTabControl1.SelectedTab = UltraTabControl1.Tabs(1)

            lblBalance.Text = txtGriege_Qty.Text
            Call Load_Grid_SockCode()
            Call Load_WithData()
            If Microsoft.VisualBasic.Left(txtQuality.Text, 1) = "Y" Then
                UltraTabControl1.Tabs(2).Enabled = True
            End If

            If Microsoft.VisualBasic.Left(txtQuality.Text, 1) = "Y" Then
            Else
                UltraTabControl1.Tabs(3).Enabled = True
            End If
        End If
    End Sub

    Function Load_Gride()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTableGrige_Stock
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
            .DisplayLayout.Bands(0).Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '  .DisplayLayout.Bands(0).Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            '.DisplayLayout.Bands(0).Columns(3).Width = 90
            '.DisplayLayout.Bands(0).Columns(4).Width = 90
            '.DisplayLayout.Bands(0).Columns(5).Width = 90
            ''  .DisplayLayout.Bands(0).Columns(6).Width = 90
            ' .DisplayLayout.Bands(0).Columns(7).Width = 90

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function

    Function Load_GrideStock()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomerSTC = CustomerDataClass.MakeDataTablePreVious_Stock
        UltraGrid2.DataSource = c_dataCustomerSTC
        With UltraGrid2
            .DisplayLayout.Bands(0).Columns(0).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 40
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(3).Width = 40
            .DisplayLayout.Bands(0).Columns(3).AutoEdit = False


            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            '  .DisplayLayout.Bands(0).Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            '.DisplayLayout.Bands(0).Columns(3).Width = 90
            '.DisplayLayout.Bands(0).Columns(4).Width = 90
            '.DisplayLayout.Bands(0).Columns(5).Width = 90
            ''  .DisplayLayout.Bands(0).Columns(6).Width = 90
            ' .DisplayLayout.Bands(0).Columns(7).Width = 90

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function

    Private Sub UltraButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton3.Click
        With UltraGroupBox9
            .Location = New Point(9, 203)
            .Width = 704
            .Height = 292
          
        End With
        With UltraGrid1
            .Width = 691
            .Height = 250
        End With
    End Sub

    Private Sub UltraButton7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton7.Click
        With UltraGroupBox9
            .Location = New Point(9, 289)
            .Width = 704
            .Height = 197
        End With

        With UltraGrid1
            .Width = 691
            .Height = 154
        End With
    End Sub

    Function Calculation_Balance()
        Dim I As Integer
        Dim Value As Double
        Dim _Vstring As String
        Try
            I = 0
            Value = 0
            For Each uRow As UltraGridRow In UltraGrid1.Rows

                With UltraGrid1
                    If IsNumeric(.Rows(I).Cells(6).Text) Then

                        If CDbl((.Rows(I).Cells(6).Text)) <= CDbl((.Rows(I).Cells(5).Text)) Then
                            Value = Value + CDbl((.Rows(I).Cells(5).Text))
                            If (.Rows(I).Cells(8).Value) = True Then
                            Else
                                .Rows(I).Cells(0).Appearance.BackColor = Color.White
                                .Rows(I).Cells(1).Appearance.BackColor = Color.White
                                .Rows(I).Cells(2).Appearance.BackColor = Color.White
                                .Rows(I).Cells(3).Appearance.BackColor = Color.White
                                .Rows(I).Cells(4).Appearance.BackColor = Color.White
                                .Rows(I).Cells(5).Appearance.BackColor = Color.White
                                .Rows(I).Cells(6).Appearance.BackColor = Color.White
                                .Rows(I).Cells(7).Appearance.BackColor = Color.White
                                .Rows(I).Cells(8).Appearance.BackColor = Color.White
                            End If
                        Else
                            MsgBox("Qty grater than to stock", MsgBoxStyle.Information, "Information ....")
                            .Rows(I).Cells(0).Appearance.BackColor = Color.Red
                            .Rows(I).Cells(1).Appearance.BackColor = Color.Red
                            .Rows(I).Cells(2).Appearance.BackColor = Color.Red
                            .Rows(I).Cells(3).Appearance.BackColor = Color.Red
                            .Rows(I).Cells(4).Appearance.BackColor = Color.Red
                            .Rows(I).Cells(5).Appearance.BackColor = Color.Red
                            .Rows(I).Cells(6).Appearance.BackColor = Color.Red
                            .Rows(I).Cells(7).Appearance.BackColor = Color.Red
                            .Rows(I).Cells(8).Appearance.BackColor = Color.Red
                            .Rows(I).Selected = True
                            Exit For
                        End If
                        '_Vstring = Value
                        '_Vstring = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        '_Vstring = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                        'UltraGrid1.Rows(I).Cells(6).Value = _Vstring
                    End If
                End With
                I = I + 1
            Next
            Value = CDbl(txtGriege_Qty.Text) - Value
            lblBalance.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            lblBalance.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            If lblBalance.Text > 0 Then
                ' UltraTabControl1.Tabs(3).Selected = True
            Else
                UltraTabControl1.Tabs(3).Enabled = False
                UltraButton6.Enabled = True
                End If

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                'Con.close()
            End If
        End Try
    End Function

   

    Private Sub UltraGrid1_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles UltraGrid1.AfterRowUpdate
        Calculation_Balance()
    End Sub

    Function Search_Tec_Spec()
        Dim i As Integer
        Dim vcWhere As String
        Dim M01 As DataSet
        Dim M02 As DataSet
        Dim Value As Double
        Dim _VString As String
        Dim Diff As TimeSpan
        Dim _To As Date
        'Dim Value As Double
        Dim _quality As String

        Try
            Dim con = New SqlConnection()
            con = DBEngin.GetConnection(True)

            If Microsoft.VisualBasic.Left(Trim(txtQuality.Text), 1) = "Q" Then
                _quality = Microsoft.VisualBasic.Right(Trim(txtQuality.Text), Microsoft.VisualBasic.Len(Trim(txtQuality.Text)) - 1)
            Else
                _quality = Trim(txtQuality.Text)
            End If
            i = 0
            vcWhere = "M22Quality='" & _quality & "' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TEC"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
            Else
                vcWhere = "M22Quality='" & Trim(txtCommon.Text) & "' "
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TEC"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                Else
                    Exit Function
                End If
            End If
            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                ' With frmYarn_Booking
                If i = 0 Then
                    vcWhere = "tmpRefNo=" & Delivary_Ref & " AND tmpDis='" & Trim(M01.Tables(0).Rows(i)("M22Yarn")) & "' "
                    dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "TYB"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(dsUser) Then
                        txtYarn1.Text = ""
                        txtCom1.Text = ""
                        txtReq1.Text = ""
                    Else

                        txtYarn1.Text = M01.Tables(0).Rows(i)("M22Yarn")
                        txtCom1.Text = CInt(M01.Tables(0).Rows(i)("M22Yarn_Cons"))
                        Value = lblBalance.Text
                        Value = Value * (txtCom1.Text / 100)
                        If IsNumeric(txtK_Wastage.Text) Then
                            Value = Value / ((100 - txtK_Wastage.Text) / 100)
                        End If
                        txtReq1.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        txtReq1.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    End If
                    '   pbCount1.Maximum = Value
                ElseIf i = 1 Then
                    vcWhere = "tmpRefNo=" & Delivary_Ref & " AND tmpDis='" & Trim(M01.Tables(0).Rows(i)("M22Yarn")) & "' "
                    dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "TYB"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(dsUser) Then
                        txtYarn2.Text = ""
                        txtCom2.Text = ""
                        txtReq2.Text = ""
                    Else
                        txtYarn2.Text = M01.Tables(0).Rows(i)("M22Yarn")
                        txtCom2.Text = CInt(M01.Tables(0).Rows(i)("M22Yarn_Cons"))

                        Value = lblBalance.Text
                        Value = Value * (txtCom2.Text / 100)
                        If IsNumeric(txtK_Wastage.Text) Then
                            Value = Value / ((100 - txtK_Wastage.Text) / 100)
                        End If
                        txtReq2.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        txtReq2.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                        '  pbCount2.Maximum = Value
                    End If
                ElseIf i = 2 Then
                    vcWhere = "tmpRefNo=" & Delivary_Ref & " AND tmpDis='" & Trim(M01.Tables(0).Rows(i)("M22Yarn")) & "' "
                    dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "TYB"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(dsUser) Then
                        txtYarn3.Text = ""
                        txtCom3.Text = ""
                        txtReq3.Text = ""
                    Else
                        txtYarn3.Text = M01.Tables(0).Rows(i)("M22Yarn")
                        txtCom3.Text = CInt(M01.Tables(0).Rows(i)("M22Yarn_Cons"))

                        Value = lblBalance.Text
                        Value = Value * (txtCom3.Text / 100)
                        If IsNumeric(txtK_Wastage.Text) Then
                            Value = Value / ((100 - txtK_Wastage.Text) / 100)
                        End If
                        txtReq3.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        txtReq3.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                        '  pbCount3.Maximum = Value
                    End If

                End If


                i = i + 1
            Next

            '----------------------------------------------------------------
            'Dim Z As Integer
            'Z = 0
            'i = 0
            'vcWhere = "M22Quality='" & Trim(frmLoad_Pln.txtQuality.Text) & "' "
            'M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TEC"), New SqlParameter("@vcWhereClause1", vcWhere))
            'For Each DTRow3 As DataRow In M01.Tables(0).Rows
            '    Z = 0
            '    vcWhere = "M33Description='" & Trim(M01.Tables(0).Rows(i)("M22Yarn")) & "'"
            '    M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TEC"), New SqlParameter("@vcWhereClause1", vcWhere))
            '    For Each DTRow4 As DataRow In M02.Tables(0).Rows
            '        Dim newRow As DataRow = c_dataCustomer1.NewRow

            '        newRow("10Class") = M02.Tables(0).Rows(Z)("M3310Class")
            '        newRow("Description") = M02.Tables(0).Rows(Z)("M33Description")
            '        newRow("Stock Code") = M02.Tables(0).Rows(Z)("M33Stock_Code")

            '        c_dataCustomer1.Rows.Add(newRow)

            '        Z = Z + 1
            '    Next
            '    i = i + 1
            'Next
            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                'Con.close()
            End If
        End Try
    End Function

    Private Sub UltraGrid1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles UltraGrid1.Click
        On Error Resume Next
        Dim _RowIndex As Integer
        _RowIndex = UltraGrid1.ActiveRow.Index
        If Trim(UltraGrid1.Rows(_RowIndex).Cells(8).Value) = True Then
            UltraGrid1.Rows(_RowIndex).Cells(6).Value = UltraGrid1.Rows(_RowIndex).Cells(5).Value
        End If
    End Sub

    Function Update_Records_Grige1()
        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Dim vcWhere As String

        Dim M01 As DataSet
        Dim i As Integer
        Dim ncQryType As String
        Try
            i = 0
            For Each uRow As UltraGridRow In UltraGrid1.Rows
                If IsNumeric(UltraGrid1.Rows(i).Cells(6).Value) Then
                    vcWhere = "T12Ref_No=" & Delivary_Ref & " and T12Sales_Order='" & strSales_Order & "' and T12Line_Item=" & strLine_Item & " and T12Stock_Code='" & Trim(UltraGrid1.Rows(i).Cells(3).Value) & "'"
                    M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetUse_Griege_Qty", New SqlParameter("@cQryType", "CGS1"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M01) Then
                    Else
                        _grgStatus = True
                        ncQryType = "GADD"
                        nvcFieldList1 = "(T12Ref_No," & "T12Sales_Order," & "T12Line_Item," & "T12Date," & "T12Time," & "T12Stock_Code," & "T12Qty," & "T12Status," & "T12Confirm_By) " & "values(" & Delivary_Ref & ",'" & strSales_Order & "'," & strLine_Item & ",'" & Today & "','" & Now & "','" & Trim(UltraGrid1.Rows(i).Cells(3).Value) & "','" & Trim(UltraGrid1.Rows(i).Cells(6).Value) & "','N','-')"
                        up_GetSetConsume_Grige(ncQryType, nvcFieldList1, vcWhere, connection, transaction)
                    End If
                End If


                i = i + 1
            Next
            If lblBalance.Text > 0 Then
            Else
                nvcFieldList1 = "update M01Sales_Order_SAP set M01Status='I' where M01Sales_Order='" & strSales_Order & "' and M01Line_Item=" & strLine_Item & ""
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            End If
            'nvcFieldList1 = "delete from tmpBlock_SalesOrder where tmpSales_Order='" & strSales_Order & "'"
            'ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            nvcFieldList1 = "update T01Delivary_Request set T01Status='C' where T01Sales_Order='" & strSales_Order & "' and T01Line_Item=" & strLine_Item & " and T01Status='A'"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            transaction.Commit()

            nvcFieldList1 = "T01Sales_Order='" & strSales_Order & "' and T01Status='A'"
            M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "T01D"), New SqlParameter("@vcWhereClause1", nvcFieldList1))
            If isValidDataset(M01) Then
                connection.Close()
                frmDelivaryQuatnew.Load_Gride_SalesOrder()
                frmDelivaryQuatnew.Load_SalesOrder()

                Me.Close()
                Exit Function
            Else
                UltraTabControl1.Tabs(3).Enabled = False
                UltraTabControl1.Tabs(4).Enabled = False
                UltraTabControl1.Tabs(5).Enabled = True
                UltraTabControl1.SelectedTab = UltraTabControl1.Tabs(5)
            End If
            connection.Close()
            ' Me.Close()
            'frmLoad_Pln.Close()
            'frmDelivaryQuatnew.Load_Gride_SalesOrder()
            'frmDelivaryQuatnew.Load_SalesOrder()

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)

                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
                connection.Close()
            End If
        End Try
    End Function

    Function Update_Records_Grige()
        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Dim vcWhere As String

        Dim M01 As DataSet
        Dim i As Integer
        Dim ncQryType As String
        Try
            i = 0
            For Each uRow As UltraGridRow In UltraGrid1.Rows
                If IsNumeric(UltraGrid1.Rows(i).Cells(5).Value) Then
                    vcWhere = "T12Ref_No=" & Delivary_Ref & " and T12Sales_Order='" & strSales_Order & "' and T12Line_Item=" & strLine_Item & " and T12Stock_Code='" & Trim(UltraGrid1.Rows(i).Cells(2).Value) & "'"
                    M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetUse_Griege_Qty", New SqlParameter("@cQryType", "CGS1"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M01) Then
                    Else
                        ncQryType = "GADD"
                        nvcFieldList1 = "(T12Ref_No," & "T12Sales_Order," & "T12Line_Item," & "T12Date," & "T12Time," & "T12Stock_Code," & "T12Qty," & "T12Status," & "T12Confirm_By) " & "values(" & Delivary_Ref & ",'" & strSales_Order & "'," & strLine_Item & ",'" & Today & "','" & Now & "','" & Trim(UltraGrid1.Rows(i).Cells(2).Value) & "','" & Trim(UltraGrid1.Rows(i).Cells(5).Value) & "','N','-')"
                        up_GetSetConsume_Grige(ncQryType, nvcFieldList1, vcWhere, connection, transaction)
                    End If
                End If


                i = i + 1
            Next

            If lblBalance.Text > 0 Then
                nvcFieldList1 = "update M01Sales_Order_SAP set M01Status='I' where M01Sales_Order='" & strSales_Order & "' and M01Line_Item=" & strLine_Item & ""
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            End If
            nvcFieldList1 = "delete from tmpBlock_SalesOrder where tmpSales_Order='" & strSales_Order & "'"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            transaction.Commit()
            connection.Close()
            'Me.Close()
            'frmLoad_Pln.Close()
            'frmDelivaryQuatnew.Load_Gride_SalesOrder()
            'frmDelivaryQuatnew.Load_SalesOrder()

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)

                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
                connection.Close()
            End If
        End Try
    End Function

    Private Sub UltraButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton6.Click
        Call Update_Records_Grige1()
    End Sub

    Private Sub UltraTabControl1_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles UltraTabControl1.SelectedTabChanged
        Dim _Status As String
        Dim _cat As String

        If UltraTabControl1.Tabs(2).Selected = True Then
            ' dg_YDP_Projection
            'Call Update_Records_Grige()
            'Call Update_Records_Grige1()
            Call Load_Detailes()
            Call Load_Gridewith_Data()
            Call Delete_Transaction()
            Call Load_Pro_YD_Plan()
            Call Load_Gride_YDPProjection()

        ElseIf UltraTabControl1.Tabs(3).Selected = True Then

            'Call Update_Records_Grige()
            'Call Update_Records_Grige1()

            Call Load_Yarn_Booking()
            Call Load_Gride_YB()
            Call Load_GrideStock_YB()
            Call Load_Grid_SockCode_YB()
            Call Load_Gridewith_Data_YB()
            '_Status = "1"
            '_cat = Microsoft.VisualBasic.Left(txtYarn1.Text, 7)
            'Call Calculation_YB_Balance(_Status, _cat)
            '_Status = "2"
            '_cat = Microsoft.VisualBasic.Left(txtYarn2.Text, 7)
            'Call Calculation_YB_Balance(_Status, _cat)
            '_Status = "3"
            '_cat = Microsoft.VisualBasic.Left(txtYarn3.Text, 7)
            'Call Calculation_YB_Balance(_Status, _cat)
            If Trim(txtShade.Text) = "D" Then
                lblYarnDis.Text = "Dark"
            ElseIf Trim(txtShade.Text) = "L" Then
                lblYarnDis.Text = "Light"

            End If
            'txtRequest_Date.Text = Today
            txtReq_Grige_YB.Text = lblBalance.Text
            txtReq_Grige_YB.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
            txtReq_Grige_YB.ReadOnly = True
            txtK_Wastage.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
            ' txtK_Wastage.ReadOnly = True

            txtCom1.ReadOnly = True
            txtCom2.ReadOnly = True
            txtCom3.ReadOnly = True

            txtCom1.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            txtCom2.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            txtCom3.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

            txtYarn1.ReadOnly = True
            txtYarn2.ReadOnly = True
            txtYarn3.ReadOnly = True

            txtYarn1.Appearance.TextHAlign = Infragistics.Win.HAlign.Left
            txtYarn2.Appearance.TextHAlign = Infragistics.Win.HAlign.Left
            txtYarn3.Appearance.TextHAlign = Infragistics.Win.HAlign.Left

            txtReq1.ReadOnly = True
            txtReq2.ReadOnly = True
            txtReq3.ReadOnly = True

            txtReq1.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            txtReq2.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            txtReq3.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

            lblBalance_YB.Text = txtReq1.Text
            lblBalance_YB1.Text = txtReq2.Text
            lblBalance_YB2.Text = txtReq3.Text

            _Status = "1"
            _cat = Microsoft.VisualBasic.Left(txtYarn1.Text, 7)
            Call Calculation_YB_Balance(_Status, _cat)
            _Status = "2"
            _cat = Microsoft.VisualBasic.Left(txtYarn2.Text, 7)
            Call Calculation_YB_Balance(_Status, _cat)
            _Status = "3"
            _cat = Microsoft.VisualBasic.Left(txtYarn3.Text, 7)
            If txtYarn1.Text <> "" Then
                lblYD_Dis1.Text = "On " & txtYarn1.Text
            End If

            If txtYarn2.Text <> "" Then
                lblYD_Dis2.Text = "On " & txtYarn2.Text
            End If

            If txtYarn3.Text <> "" Then
                lblYD_Dis3.Text = "On " & txtYarn3.Text
            End If
            Call Calculation_YB_Balance(_Status, _cat)

        ElseIf UltraTabControl1.Tabs(4).Selected = True Then

            'Call Update_Records_Grige()
            'Call Update_Records_Grige1()
            txtSales_Order.Text = strSales_Order
            txtLine_Knt.Text = strLine_Item
            txtQuality_knt.Text = txtQuality.Text
            If txtReq_Grige_YB.Text <> "" Then
                strQty = txtReq_Grige_YB.Text
            Else
                strQty = txtReq_Greige.Text
            End If
            txtQty_Knt.Text = (strQty.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            txtQty_Knt.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", strQty))
            lblKnt_Balance.Text = txtQty_Knt.Text
            ' strPer_Day = strQty / 24
            txtDaily_Capacity.Text = (strPer_Day.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            txtDaily_Capacity.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", strPer_Day))

            If txtQty_Knt.Text <> "" And txtDaily_Capacity.Text <> "" Then
                txtDays.Text = CInt(txtQty_Knt.Text / txtDaily_Capacity.Text)
            End If

            txtMC_Group_Knt.Text = strMC_Group
            Call Load_Gride_Knt()
            Call Quality_Group()
            Call Load_Gride_KNTProjection()
            Call Load_Pro_KNT_Week()

            txtK_Qty.Text = txtQty_Knt.Text
            'Call Load_Gride_projection()
            'Call Load_Projection_Detailes(txtDelivary_Date.Text)
        ElseIf UltraTabControl1.Tabs(5).Selected = True Then

            Call Load_Gride_Dye_Grige() 'Screen 1 Gride Creation
            Call Load_Gride_Dye_Projection()

            Call Load_Gride_Dye_Main() 'Screen 1 Data Filling

            Call Load_Gride_Dye_Grige_Detailes() 'Screen 2 Creation
            Call Load_Gride_Dye_Capacity() 'Dye Capacity
            Call Dye_Shade_Gride_Main()

            Call Load_Gride_Delivary()

            txtDye_Bulk.ReadOnly = True
            txtDye_BulkApp.ReadOnly = True
            txtDye_NPL.ReadOnly = True
            txtDye_NPL_App.ReadOnly = True
            txtDye_PP.ReadOnly = True
            txtDye_LDApp.ReadOnly = True
            txtDye_LD.ReadOnly = True
            txtDye_LDApp.ReadOnly = True
            txtDye_PP_App.ReadOnly = True

            txtDye_Week.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            txtDye_Qty.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

            txtDye_Bulk.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            txtDye_BulkApp.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            txtDye_NPL.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            txtDye_NPL_App.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            txtDye_PP.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            txtDye_LDApp.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            txtDye_LD.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            txtDye_LDApp.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            txtDye_PP_App.Appearance.TextHAlign = Infragistics.Win.HAlign.Center


        ElseIf UltraTabControl1.Tabs(6).Selected = True Then
            chkN_Lead.Checked = True
            Call Load_Grid_Delivary_Plan()
            'Call Load_Gride()
            'Call Load_GrideStock()
            'Call Load_Grid_SockCode()
            'Call Load_WithData()
        End If

    End Sub


    Public Function weekNumber(ByVal d As Date) As Integer
        weekNumber = DatePart(DateInterval.WeekOfYear, d, FirstDayOfWeek.Monday, FirstWeekOfYear.System)

    End Function

    Function Load_Pro_KNT_Week()
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

        Dim agroup1_1 As UltraGridGroup
        Dim agroup1_2 As UltraGridGroup
        Dim agroup1_3 As UltraGridGroup
        Dim agroup1_4 As UltraGridGroup
        Dim agroup1_5 As UltraGridGroup

        Dim agroup2_1 As UltraGridGroup
        Dim agroup2_2 As UltraGridGroup
        Dim agroup2_3 As UltraGridGroup
        Dim agroup2_4 As UltraGridGroup
        Dim agroup2_5 As UltraGridGroup

        Dim agroup3_1 As UltraGridGroup
        Dim agroup3_2 As UltraGridGroup
        Dim agroup3_3 As UltraGridGroup
        Dim agroup3_4 As UltraGridGroup
        Dim agroup3_5 As UltraGridGroup

        Dim agroup4_1 As UltraGridGroup
        Dim agroup4_2 As UltraGridGroup
        Dim agroup4_3 As UltraGridGroup
        Dim agroup4_4 As UltraGridGroup
        Dim agroup4_5 As UltraGridGroup

        Dim agroup5_1 As UltraGridGroup
        Dim agroup5_2 As UltraGridGroup
        Dim agroup5_3 As UltraGridGroup
        Dim agroup5_4 As UltraGridGroup
        Dim agroup5_5 As UltraGridGroup

        Dim _Date As Date
        Dim countdays As Integer
        Dim _WeekNo As Integer
        Dim _FromDate As Date
        Dim _ColumCount As Integer
        Dim Value As Double
        Dim Value4 As Double
        Dim _STSting As String


        Try
            'Dim agroup1 As UltraGridGroup
            'Dim agroup2 As UltraGridGroup
            'Dim agroup3 As UltraGridGroup
            'Dim agroup4 As UltraGridGroup
            'Dim agroup5 As UltraGridGroup
            '  Dim agroup6 As UltraGridGroup
            dg_Knt_Week1.DisplayLayout.Bands(0).Groups.Clear()
            dg_Knt_Week1.DisplayLayout.Bands(0).Columns.Dispose()

            'If UltraGrid3.DisplayLayout.Bands(0).GroupHeadersVisible = True Then
            'Else
            '  agroup1.Key = ""
            '  agroup1 = UltraGrid3.DisplayLayout.Bands(0).Groups.Remove("GroupH")
            agroup1 = dg_Knt_Week1.DisplayLayout.Bands(0).Groups.Add("GroupH")
            'Me.UltraGrid3.DisplayLayout.Bands(0).Columns.Add("Line", "Line Item")
            'Me.UltraGrid3.DisplayLayout.Bands(0).Columns("Line").Group = agroup1
            'Me.UltraGrid3.DisplayLayout.Bands(0).Columns("Line").Width = 50

            'Me.UltraGrid3.DisplayLayout.Bands(0).Columns.Add("##", "##")
            'Me.UltraGrid3.DisplayLayout.Bands(0).Columns("##").Group = agroup1
            'Me.UltraGrid3.DisplayLayout.Bands(0).Columns("##").Width = 120
            ''  End If
            ' agroup1 = UltraGrid3.DisplayLayout.Bands(0).Groups.Remove(0)


            _Code = _Code - 1
            agroup1.Header.Caption = ""

            agroup1.Width = 110
            Dim dt As DataTable = New DataTable()
            ' dt.Columns.Add("ID", GetType(Integer))
            Dim colWork As New DataColumn("##", GetType(String))
            dt.Columns.Add(colWork)
            colWork.ReadOnly = True


            dt.Rows.Add("Confimed Projection")
            dt.Rows.Add("Consumed Projection")
            dt.Rows.Add("Balance Projection")
            dt.Rows.Add("Projection (" & strDisname & ")")
            dt.Rows.Add("Consumed Projection(" & strDisname & ")")
            dt.Rows.Add("Balance Projection(" & strDisname & ")")
            dt.Rows.Add("")
            dt.Rows.Add("Allocated Projection ")
            dt.Rows.Add("Consumed Projection")
            dt.Rows.Add("Balance Projection")



            Me.dg_Knt_Week1.SetDataBinding(dt, Nothing)
            Me.dg_Knt_Week1.DisplayLayout.Bands(0).Columns(0).Group = agroup1
            'Me.dg_Knt_Week.DisplayLayout.Bands(0).Columns(1).Group = agroup1
            'Me.dg_Knt_Week.DisplayLayout.Bands(0).Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            Me.dg_Knt_Week1.DisplayLayout.Bands(0).Columns(0).Width = 180
            ' Me.dg_Knt_Week.DisplayLayout.Bands(0).Columns(1).Width = 50
            Dim _Group As String
            'agroup2.Key = ""
            'agroup3.Key = ""
            'agroup4.Key = ""
            '' agroup5.Key = ""
           

            'Knitting
            I = 0
            vcWhere = "T15Sales_Order='" & strSales_Order & "' and T15Line_Item=" & txtLine_Item.Text & ""
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "PROX"), New SqlParameter("@vcWhereClause1", vcWhere))
            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                Dim remain As Integer
                Dim noOfWeek As Integer
                Dim userDate As Date
                Dim _LastDate As Date
                Dim _TimeSpan As TimeSpan

                userDate = DateTime.Parse(M01.Tables(0).Rows(I)("T15Month") & "/1/" & M01.Tables(0).Rows(I)("T15Year"))
                ' MsgBox(WeekdayName(Weekday(userDate)))
                If WeekdayName(Weekday(userDate)) = "Sunday" Then
                    userDate = userDate.AddDays(-3)
                ElseIf WeekdayName(Weekday(userDate)) = "Monday" Then
                    userDate = userDate.AddDays(-4)
                ElseIf WeekdayName(Weekday(userDate)) = "Tuesday" Then
                    userDate = userDate.AddDays(-5)
                ElseIf WeekdayName(Weekday(userDate)) = "Thusday" Then
                    'userDate = userDate.AddDays(-1)
                ElseIf WeekdayName(Weekday(userDate)) = "Friday" Then
                    userDate = userDate.AddDays(-1)
                ElseIf WeekdayName(Weekday(userDate)) = "Saturday" Then
                    userDate = userDate.AddDays(-2)

                End If


                _LastDate = DateTime.Parse(M01.Tables(0).Rows(I)("T15Month") & "/1/" & M01.Tables(0).Rows(I)("T15Year"))
                ' MsgBox(Date.DaysInMonth(_LastDate.Year, _LastDate.Month))
                _LastDate = DateTime.Parse(M01.Tables(0).Rows(I)("T15Month") & "/" & Date.DaysInMonth(_LastDate.Year, _LastDate.Month) & "/" & M01.Tables(0).Rows(I)("T15Year"))
                If WeekdayName(Weekday(_LastDate)) = "Sunday" Then
                    _LastDate = _LastDate.AddDays(-4)
                ElseIf WeekdayName(Weekday(_LastDate)) = "Monday" Then
                    _LastDate = _LastDate.AddDays(-5)
                ElseIf WeekdayName(Weekday(_LastDate)) = "Tuesday" Then
                    _LastDate = _LastDate.AddDays(-6)
                ElseIf WeekdayName(Weekday(_LastDate)) = "Thusday" Then
                    'userDate = userDate.AddDays(-1)
                ElseIf WeekdayName(Weekday(_LastDate)) = "Friday" Then
                    _LastDate = _LastDate.AddDays(-2)
                ElseIf WeekdayName(Weekday(_LastDate)) = "Saturday" Then
                    _LastDate = _LastDate.AddDays(-3)

                End If

                _TimeSpan = _LastDate.Subtract(userDate)
                _WeekNo = _TimeSpan.Days / 7
                I = I + 1
            Next

            agroup3 = dg_Knt_Week1.DisplayLayout.Bands(0).Groups.Add("Group3")
            agroup3.Header.Caption = "Knitting"
            _WeekNo = _WeekNo * 60
            agroup3.Width = _WeekNo
            '=====================================================================================
            I = 0
            vcWhere = "T15Sales_Order='" & strSales_Order & "' and T15Line_Item=" & txtLine_Item.Text & " "
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "PROX"), New SqlParameter("@vcWhereClause1", vcWhere))
            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                Dim remain As Integer
                Dim noOfWeek As Integer
                Dim userDate As Date
                Dim _LastDate As Date
                Dim _TimeSpan As TimeSpan

                userDate = DateTime.Parse(M01.Tables(0).Rows(I)("T15Month") & "/1/" & M01.Tables(0).Rows(I)("T15Year"))
                ' MsgBox(WeekdayName(Weekday(userDate)))
                If WeekdayName(Weekday(userDate)) = "Sunday" Then
                    userDate = userDate.AddDays(-3)
                ElseIf WeekdayName(Weekday(userDate)) = "Monday" Then
                    userDate = userDate.AddDays(-4)
                ElseIf WeekdayName(Weekday(userDate)) = "Tuesday" Then
                    userDate = userDate.AddDays(-5)
                ElseIf WeekdayName(Weekday(userDate)) = "Thusday" Then
                    'userDate = userDate.AddDays(-1)
                ElseIf WeekdayName(Weekday(userDate)) = "Friday" Then
                    userDate = userDate.AddDays(-1)
                ElseIf WeekdayName(Weekday(userDate)) = "Saturday" Then
                    userDate = userDate.AddDays(-2)

                End If
                _LastDate = DateTime.Parse(M01.Tables(0).Rows(I)("T15Month") & "/1/" & M01.Tables(0).Rows(I)("T15Year"))
                ' MsgBox(Date.DaysInMonth(_LastDate.Year, _LastDate.Month))
                _LastDate = DateTime.Parse(M01.Tables(0).Rows(I)("T15Month") & "/" & Date.DaysInMonth(_LastDate.Year, _LastDate.Month) & "/" & M01.Tables(0).Rows(I)("T15Year"))

                If WeekdayName(Weekday(_LastDate)) = "Sunday" Then
                    _LastDate = _LastDate.AddDays(-4)
                ElseIf WeekdayName(Weekday(_LastDate)) = "Monday" Then
                    _LastDate = _LastDate.AddDays(-5)
                ElseIf WeekdayName(Weekday(_LastDate)) = "Tuesday" Then
                    _LastDate = _LastDate.AddDays(-6)
                ElseIf WeekdayName(Weekday(_LastDate)) = "Thusday" Then
                    'userDate = userDate.AddDays(-1)
                ElseIf WeekdayName(Weekday(_LastDate)) = "Friday" Then
                    _LastDate = _LastDate.AddDays(-2)
                ElseIf WeekdayName(Weekday(_LastDate)) = "Saturday" Then
                    _LastDate = _LastDate.AddDays(-3)

                End If

                _TimeSpan = _LastDate.Subtract(userDate)
                _WeekNo = _TimeSpan.Days / 7

                userDate = userDate.AddDays(+7)
                vcWhere = "T15Sales_Order='" & Trim(cboSO.Text) & "' AND t01bulk ='1st Bulk' and T15Year=" & M01.Tables(0).Rows(I)("T15Year") & " and T15Month=" & M01.Tables(0).Rows(I)("T15Month") & ""
                T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "FSTB"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(T01) Then
                    userDate = userDate.AddDays(-21)
                Else
                    vcWhere = "T01Sales_Order='" & Trim(cboSO.Text) & "' AND T01Line_Item  ='" & strLine_Item & "'"
                    T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "PRNT"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(T01) Then
                        userDate = userDate.AddDays(-28)
                    Else
                        userDate = userDate.AddDays(-21)
                    End If
                    End If

                    If _WeekNo = 5 Then
                        Dim culture As System.Globalization.CultureInfo
                        Dim intWeek As Integer
                        Dim _StrWeek As String
                        Dim _StrWeek1 As String

                        culture = System.Globalization.CultureInfo.CurrentCulture
                    intWeek = culture.Calendar.GetWeekOfYear(userDate, System.Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Thursday)

                        _StrWeek = "Week " & intWeek
                        _StrWeek1 = "Week " & intWeek & I
                        Me.dg_Knt_Week1.DisplayLayout.Bands(0).Columns.Add(_StrWeek1, _StrWeek)
                        Me.dg_Knt_Week1.DisplayLayout.Bands(0).Columns(_StrWeek1).Group = agroup3
                        Me.dg_Knt_Week1.DisplayLayout.Bands(0).Columns(_StrWeek1).Width = 60
                        Me.dg_Knt_Week1.DisplayLayout.Bands(0).Columns(_StrWeek1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                        userDate = userDate.AddDays(+7)
                        culture = System.Globalization.CultureInfo.CurrentCulture
                    intWeek = culture.Calendar.GetWeekOfYear(userDate, System.Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Thursday)

                        _StrWeek = "Week " & intWeek
                        _StrWeek1 = "Week " & intWeek & I

                        Me.dg_Knt_Week1.DisplayLayout.Bands(0).Columns.Add(_StrWeek1, _StrWeek)
                        Me.dg_Knt_Week1.DisplayLayout.Bands(0).Columns(_StrWeek1).Group = agroup3
                        Me.dg_Knt_Week1.DisplayLayout.Bands(0).Columns(_StrWeek1).Width = 60
                        Me.dg_Knt_Week1.DisplayLayout.Bands(0).Columns(_StrWeek1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                        userDate = userDate.AddDays(+7)
                        culture = System.Globalization.CultureInfo.CurrentCulture
                    intWeek = culture.Calendar.GetWeekOfYear(userDate, System.Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Thursday)

                        _StrWeek = "Week " & intWeek
                        _StrWeek1 = "Week " & intWeek & I
                        Me.dg_Knt_Week1.DisplayLayout.Bands(0).Columns.Add(_StrWeek1, _StrWeek)
                        Me.dg_Knt_Week1.DisplayLayout.Bands(0).Columns(_StrWeek1).Group = agroup3
                        Me.dg_Knt_Week1.DisplayLayout.Bands(0).Columns(_StrWeek1).Width = 60
                        Me.dg_Knt_Week1.DisplayLayout.Bands(0).Columns(_StrWeek1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                        userDate = userDate.AddDays(+7)
                        culture = System.Globalization.CultureInfo.CurrentCulture
                    intWeek = culture.Calendar.GetWeekOfYear(userDate, System.Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Thursday)

                        _StrWeek = "Week " & intWeek
                        _StrWeek1 = "Week " & intWeek & I
                        Me.dg_Knt_Week1.DisplayLayout.Bands(0).Columns.Add(_StrWeek1, _StrWeek)
                        Me.dg_Knt_Week1.DisplayLayout.Bands(0).Columns(_StrWeek1).Group = agroup3
                        Me.dg_Knt_Week1.DisplayLayout.Bands(0).Columns(_StrWeek1).Width = 60
                        Me.dg_Knt_Week1.DisplayLayout.Bands(0).Columns(_StrWeek1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                        userDate = userDate.AddDays(+7)
                        culture = System.Globalization.CultureInfo.CurrentCulture
                    intWeek = culture.Calendar.GetWeekOfYear(userDate, System.Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Thursday)

                        _StrWeek = "Week " & intWeek
                        _StrWeek1 = "Week " & intWeek & I

                        Me.dg_Knt_Week1.DisplayLayout.Bands(0).Columns.Add(_StrWeek1, _StrWeek)
                        Me.dg_Knt_Week1.DisplayLayout.Bands(0).Columns(_StrWeek1).Group = agroup3
                        Me.dg_Knt_Week1.DisplayLayout.Bands(0).Columns(_StrWeek1).Width = 60
                        Me.dg_Knt_Week1.DisplayLayout.Bands(0).Columns(_StrWeek1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                    ElseIf _WeekNo = 4 Then
                        Dim culture As System.Globalization.CultureInfo
                        Dim intWeek As Integer
                        Dim _StrWeek As String

                        culture = System.Globalization.CultureInfo.CurrentCulture
                    intWeek = culture.Calendar.GetWeekOfYear(userDate, System.Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Thursday)

                        _StrWeek = "Week " & intWeek
                        Me.dg_Knt_Week1.DisplayLayout.Bands(0).Columns.Add(_StrWeek, _StrWeek)
                        Me.dg_Knt_Week1.DisplayLayout.Bands(0).Columns(_StrWeek).Group = agroup3
                        Me.dg_Knt_Week1.DisplayLayout.Bands(0).Columns(_StrWeek).Width = 60
                        Me.dg_Knt_Week1.DisplayLayout.Bands(0).Columns(_StrWeek).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                        userDate = userDate.AddDays(+7)
                        culture = System.Globalization.CultureInfo.CurrentCulture
                    intWeek = culture.Calendar.GetWeekOfYear(userDate, System.Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Thursday)

                        _StrWeek = "Week " & intWeek
                        Me.dg_Knt_Week1.DisplayLayout.Bands(0).Columns.Add(_StrWeek, _StrWeek)
                        Me.dg_Knt_Week1.DisplayLayout.Bands(0).Columns(_StrWeek).Group = agroup3
                        Me.dg_Knt_Week1.DisplayLayout.Bands(0).Columns(_StrWeek).Width = 60
                        Me.dg_Knt_Week1.DisplayLayout.Bands(0).Columns(_StrWeek).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                        userDate = userDate.AddDays(+7)
                        culture = System.Globalization.CultureInfo.CurrentCulture
                    intWeek = culture.Calendar.GetWeekOfYear(userDate, System.Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Thursday)

                        _StrWeek = "Week " & intWeek
                        Me.dg_Knt_Week1.DisplayLayout.Bands(0).Columns.Add(_StrWeek, _StrWeek)
                        Me.dg_Knt_Week1.DisplayLayout.Bands(0).Columns(_StrWeek).Group = agroup3
                        Me.dg_Knt_Week1.DisplayLayout.Bands(0).Columns(_StrWeek).Width = 60
                        Me.dg_Knt_Week1.DisplayLayout.Bands(0).Columns(_StrWeek).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                        userDate = userDate.AddDays(+7)
                        culture = System.Globalization.CultureInfo.CurrentCulture
                    intWeek = culture.Calendar.GetWeekOfYear(userDate, System.Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Thursday)

                        _StrWeek = "Week " & intWeek
                        Me.dg_Knt_Week1.DisplayLayout.Bands(0).Columns.Add(_StrWeek, _StrWeek)
                        Me.dg_Knt_Week1.DisplayLayout.Bands(0).Columns(_StrWeek).Group = agroup3
                        Me.dg_Knt_Week1.DisplayLayout.Bands(0).Columns(_StrWeek).Width = 60
                        Me.dg_Knt_Week1.DisplayLayout.Bands(0).Columns(_StrWeek).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                    End If
                    I = I + 1
            Next
            '================================================================================>>
            'DATA FILLING
            'CONFIRM PROJECTION
            Dim _ROWCOUNT As Integer
            Dim Z As Integer
            Dim _QUALITY As String
            Dim Y As Integer

            _QUALITY = ""
            vcWhere = "select * from P01PARAMETER where P01CODE='PRN'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, vcWhere)
            If isValidDataset(M01) Then
                _Code = M01.Tables(0).Rows(0)("P01NO")
            End If

            _Code = _Code - 1

            _Rowindex = 0
            _COLUMCOUNT = 1

            I = 0
            vcWhere = "T15Sales_Order='" & strSales_Order & "' AND T15Line_Item=" & strLine_Item & ""
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "PROX"), New SqlParameter("@vcWhereClause1", vcWhere))
            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                Dim remain As Integer
                Dim noOfWeek As Integer
                Dim userDate As Date
                Dim _LastDate As Date
                Dim _TimeSpan As TimeSpan
                Dim _St As String

                userDate = DateTime.Parse(M01.Tables(0).Rows(I)("T15Month") & "/1/" & M01.Tables(0).Rows(I)("T15Year"))
                ' MsgBox(WeekdayName(Weekday(userDate)))
                If WeekdayName(Weekday(userDate)) = "Sunday" Then
                    userDate = userDate.AddDays(-3)
                ElseIf WeekdayName(Weekday(userDate)) = "Monday" Then
                    userDate = userDate.AddDays(-4)
                ElseIf WeekdayName(Weekday(userDate)) = "Tuesday" Then
                    userDate = userDate.AddDays(-5)
                ElseIf WeekdayName(Weekday(userDate)) = "Thusday" Then
                    'userDate = userDate.AddDays(-1)
                ElseIf WeekdayName(Weekday(userDate)) = "Friday" Then
                    userDate = userDate.AddDays(-1)
                ElseIf WeekdayName(Weekday(userDate)) = "Saturday" Then
                    userDate = userDate.AddDays(-2)

                End If

                _LastDate = DateTime.Parse(M01.Tables(0).Rows(I)("T15Month") & "/1/" & M01.Tables(0).Rows(I)("T15Year"))
                ' MsgBox(Date.DaysInMonth(_LastDate.Year, _LastDate.Month))
                _LastDate = DateTime.Parse(M01.Tables(0).Rows(I)("T15Month") & "/" & Date.DaysInMonth(_LastDate.Year, _LastDate.Month) & "/" & M01.Tables(0).Rows(I)("T15Year"))
                If WeekdayName(Weekday(_LastDate)) = "Sunday" Then
                    _LastDate = _LastDate.AddDays(-4)
                ElseIf WeekdayName(Weekday(_LastDate)) = "Monday" Then
                    _LastDate = _LastDate.AddDays(-5)
                ElseIf WeekdayName(Weekday(_LastDate)) = "Tuesday" Then
                    _LastDate = _LastDate.AddDays(-6)
                ElseIf WeekdayName(Weekday(_LastDate)) = "Thusday" Then
                    'userDate = userDate.AddDays(-1)
                ElseIf WeekdayName(Weekday(_LastDate)) = "Friday" Then
                    _LastDate = _LastDate.AddDays(-2)
                ElseIf WeekdayName(Weekday(_LastDate)) = "Saturday" Then
                    _LastDate = _LastDate.AddDays(-3)

                End If

                _TimeSpan = _LastDate.Subtract(userDate)
                _WeekNo = _TimeSpan.Days / 7

                Y = 0
                vcWhere = "T15Sales_Order='" & strSales_Order & "' AND T15Line_Item=" & strLine_Item & ""
                T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "QPJQ"), New SqlParameter("@vcWhereClause1", vcWhere))
                For Each DTRow4 As DataRow In T01.Tables(0).Rows
                    If Y = 0 Then
                        _QUALITY = T01.Tables(0).Rows(Y)("T15Quality")
                    Else
                        _QUALITY = _QUALITY & "','" & T01.Tables(0).Rows(Y)("T15Quality")
                    End If
                    Y = Y + 1
                Next
                vcWhere = "M43Year=" & M01.Tables(0).Rows(I)("T15Year") & " AND M43Product_Month=" & M01.Tables(0).Rows(I)("T15Month") & " AND M22Fabric_Type='" & txtFabric.Text & "' AND M22Strich_Lenth=0 AND M43Planned='IH Full Production'"
                dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "MPRO"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(dsUser) Then
                    Value = dsUser.Tables(0).Rows(0)("M43Sales_Volume") / _WeekNo
                    _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                    'userDate = userDate.AddDays(+7)
                    'vcWhere = "T15Sales_Order='" & Trim(cboSO.Text) & "' AND t01bulk='1st Bulk' and T15Year=" & T01.Tables(0).Rows(i)("T15Year") & " and T15Month=" & T01.Tables(0).Rows(i)("T15Month") & ""
                    'T02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "FSTB"), New SqlParameter("@vcWhereClause1", vcWhere))
                    'If isValidDataset(T02) Then
                    '    userDate = userDate.AddDays(-21)
                    'Else
                    '    userDate = userDate.AddDays(-14)
                    'End If

                    For Z = 1 To _WeekNo
                        If Trim(dg_Knt_Week1.Rows(_ROWCOUNT).Cells(_ColumCount).Text) = "-" Then
                        Else
                            'Dim culture As System.Globalization.CultureInfo
                            'Dim intWeek As Integer
                            'Dim _StrWeek As String

                            'culture = System.Globalization.CultureInfo.CurrentCulture
                            'intWeek = culture.Calendar.GetWeekOfYear(userDate, System.Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday)

                            dg_Knt_Week1.Rows(_ROWCOUNT).Cells(_ColumCount).Value = _St
                            dg_Knt_Week1.Rows(2).Cells(_ColumCount).Appearance.BackColor = Color.Yellow
                            dg_Knt_Week1.Rows(2).Cells(0).Appearance.BackColor = Color.Yellow

                            dg_Knt_Week1.Rows(5).Cells(_ColumCount).Appearance.BackColor = Color.LightBlue
                            dg_Knt_Week1.Rows(5).Cells(0).Appearance.BackColor = Color.LightBlue

                            userDate = userDate.AddDays(+7)
                            _ColumCount = _ColumCount + 1
                        End If
                    Next
                Else
                    _ColumCount = 1
                    For Z = 1 To _WeekNo
                        If Trim(dg_Knt_Week1.Rows(_ROWCOUNT).Cells(_ColumCount).Text) = "-" Then
                        Else
                            dg_Knt_Week1.Rows(_ROWCOUNT).Cells(_ColumCount).Value = "-"
                            'UltraGrid2.Rows(_ROWCOUNT + 1).Cells(_ColumCount).Value = "-"
                            dg_Knt_Week1.Rows(2).Cells(_ColumCount).Appearance.BackColor = Color.Yellow
                            dg_Knt_Week1.Rows(2).Cells(0).Appearance.BackColor = Color.Yellow

                            dg_Knt_Week1.Rows(5).Cells(_ColumCount).Appearance.BackColor = Color.LightBlue
                            dg_Knt_Week1.Rows(5).Cells(0).Appearance.BackColor = Color.LightBlue

                            _ColumCount = _ColumCount + 1

                        End If
                    Next
                End If
                'PLANNERS PROJECTION
                _ColumCount = 1
                _ROWCOUNT = _ROWCOUNT + 3
                vcWhere = "M43Year=" & M01.Tables(0).Rows(I)("T15Year") & " AND M43Product_Month=" & M01.Tables(0).Rows(I)("T15Month") & " AND M22Fabric_Type='" & txtFabric.Text & "' AND M22Strich_Lenth=0 AND M43Planned='IH Full Production' AND M43Planner='" & strDisname & "'"
                dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "MPRO"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(dsUser) Then
                    Value = dsUser.Tables(0).Rows(0)("M43Sales_Volume") / _WeekNo
                    _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                    'userDate = userDate.AddDays(+7)
                    'vcWhere = "T15Sales_Order='" & Trim(cboSO.Text) & "' AND t01bulk='1st Bulk' and T15Year=" & T01.Tables(0).Rows(i)("T15Year") & " and T15Month=" & T01.Tables(0).Rows(i)("T15Month") & ""
                    'T02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "FSTB"), New SqlParameter("@vcWhereClause1", vcWhere))
                    'If isValidDataset(T02) Then
                    '    userDate = userDate.AddDays(-21)
                    'Else
                    '    userDate = userDate.AddDays(-14)
                    'End If

                    For Z = 1 To _WeekNo
                        If Trim(dg_Knt_Week1.Rows(_ROWCOUNT).Cells(_ColumCount).Text) = "-" Then
                        Else
                            'Dim culture As System.Globalization.CultureInfo
                            'Dim intWeek As Integer
                            'Dim _StrWeek As String

                            'culture = System.Globalization.CultureInfo.CurrentCulture
                            'intWeek = culture.Calendar.GetWeekOfYear(userDate, System.Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday)

                            dg_Knt_Week1.Rows(_ROWCOUNT).Cells(_ColumCount).Value = _St
                            userDate = userDate.AddDays(+7)
                            _ColumCount = _ColumCount + 1
                        End If
                    Next
                Else
                    For Z = 1 To _WeekNo
                        If Trim(dg_Knt_Week1.Rows(_ROWCOUNT).Cells(_ColumCount).Text) = "-" Then
                        Else
                            dg_Knt_Week1.Rows(_ROWCOUNT).Cells(_ColumCount).Value = "-"
                            'UltraGrid2.Rows(_ROWCOUNT + 1).Cells(_ColumCount).Value = "-"
                            _ColumCount = _ColumCount + 1
                        End If
                    Next
                End If

                I = I + 1
            Next
            '===============================================================================
            'CONSUME PROJECTION
            I = 0
            _ROWCOUNT = 1
            _ColumCount = 1
            vcWhere = "T15Sales_Order='" & strSales_Order & "' and T15Line_Item=" & txtLine_Item.Text & ""
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "PROX"), New SqlParameter("@vcWhereClause1", vcWhere))
            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                Dim remain As Integer
                Dim noOfWeek As Integer
                Dim userDate As Date
                Dim _LastDate As Date
                Dim _TimeSpan As TimeSpan

                userDate = DateTime.Parse(M01.Tables(0).Rows(I)("T15Month") & "/1/" & M01.Tables(0).Rows(I)("T15Year"))
                ' MsgBox(WeekdayName(Weekday(userDate)))
                If WeekdayName(Weekday(userDate)) = "Sunday" Then
                    userDate = userDate.AddDays(-3)
                ElseIf WeekdayName(Weekday(userDate)) = "Monday" Then
                    userDate = userDate.AddDays(-4)
                ElseIf WeekdayName(Weekday(userDate)) = "Tuesday" Then
                    userDate = userDate.AddDays(-5)
                ElseIf WeekdayName(Weekday(userDate)) = "Thusday" Then
                    'userDate = userDate.AddDays(-1)
                ElseIf WeekdayName(Weekday(userDate)) = "Friday" Then
                    userDate = userDate.AddDays(-1)
                ElseIf WeekdayName(Weekday(userDate)) = "Saturday" Then
                    userDate = userDate.AddDays(-2)

                End If
                _LastDate = DateTime.Parse(M01.Tables(0).Rows(I)("T15Month") & "/1/" & M01.Tables(0).Rows(I)("T15Year"))
                ' MsgBox(Date.DaysInMonth(_LastDate.Year, _LastDate.Month))
                _LastDate = DateTime.Parse(M01.Tables(0).Rows(I)("T15Month") & "/" & Date.DaysInMonth(_LastDate.Year, _LastDate.Month) & "/" & M01.Tables(0).Rows(I)("T15Year"))
                If WeekdayName(Weekday(_LastDate)) = "Sunday" Then
                    _LastDate = _LastDate.AddDays(-4)
                ElseIf WeekdayName(Weekday(_LastDate)) = "Monday" Then
                    _LastDate = _LastDate.AddDays(-5)
                ElseIf WeekdayName(Weekday(_LastDate)) = "Tuesday" Then
                    _LastDate = _LastDate.AddDays(-6)
                ElseIf WeekdayName(Weekday(_LastDate)) = "Thusday" Then
                    'userDate = userDate.AddDays(-1)
                ElseIf WeekdayName(Weekday(_LastDate)) = "Friday" Then
                    _LastDate = _LastDate.AddDays(-2)
                ElseIf WeekdayName(Weekday(_LastDate)) = "Saturday" Then
                    _LastDate = _LastDate.AddDays(-3)

                End If

                _TimeSpan = _LastDate.Subtract(userDate)
                _WeekNo = _TimeSpan.Days / 7

                userDate = userDate.AddDays(+7)
                vcWhere = "T15Sales_Order='" & Trim(cboSO.Text) & "' AND t01bulk ='1st Bulk' and T15Year=" & M01.Tables(0).Rows(I)("T15Year") & " and T15Month=" & M01.Tables(0).Rows(I)("T15Month") & ""
                T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "FSTB"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(T01) Then
                    userDate = userDate.AddDays(-21)
                Else
                    vcWhere = "T01Sales_Order='" & Trim(cboSO.Text) & "' AND T01Line_Item  ='" & strLine_Item & "'"
                    T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "PRNT"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(T01) Then
                        userDate = userDate.AddDays(-21)
                    Else
                        userDate = userDate.AddDays(-14)
                    End If
                End If

                _QUALITY = ""
                Y = 0
                vcWhere = "T15Sales_Order='" & strSales_Order & "' AND T15Line_Item=" & strLine_Item & ""
                T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "QPJQ"), New SqlParameter("@vcWhereClause1", vcWhere))
                For Each DTRow4 As DataRow In T01.Tables(0).Rows
                    If Y = 0 Then
                        _QUALITY = T01.Tables(0).Rows(Y)("T15Quality")
                    Else
                        _QUALITY = _QUALITY & "','" & T01.Tables(0).Rows(Y)("T15Quality")
                    End If
                    Y = Y + 1
                Next

                For Z = 1 To _WeekNo
                    Dim culture As System.Globalization.CultureInfo
                    Dim intWeek As Integer
                    Dim _StrWeek As String
                    Dim _St As String
                    culture = System.Globalization.CultureInfo.CurrentCulture
                    intWeek = culture.Calendar.GetWeekOfYear(userDate, System.Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Thursday)

                    vcWhere = "tmpQuality IN ('" & _QUALITY & "') AND tmpYear=" & Year(userDate) & " AND tmpWeek_No=" & intWeek & ""
                    T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "CKNG"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(T01) Then
                        Value = T01.Tables(0).Rows(0)("QTY")
                        _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                        dg_Knt_Week1.Rows(_ROWCOUNT).Cells(_ColumCount).Value = _St
                        'UltraGrid2.Rows(_ROWCOUNT + 1).Cells(_ColumCount).Value = "-"
                        vcWhere = "tmpQuality IN ('" & _QUALITY & "') AND tmpYear=" & Year(userDate) & " AND tmpWeek_No=" & intWeek & " and tmpUser='" & strDisname & "'"
                        T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "CKNG"), New SqlParameter("@vcWhereClause1", vcWhere))
                        If isValidDataset(T01) Then
                            Value = T01.Tables(0).Rows(0)("QTY")
                            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                            dg_Knt_Week1.Rows(_ROWCOUNT).Cells(_ColumCount).Value = _St
                            dg_Knt_Week1.Rows(_ROWCOUNT + 1).Cells(_ColumCount).Value = dg_Knt_Week1.Rows(_ROWCOUNT - 1).Cells(_ColumCount).Value - dg_Knt_Week1.Rows(_ROWCOUNT).Cells(_ColumCount).Value
                            ' _ColumCount = _ColumCount + 1

                        Else
                            If Trim(dg_Knt_Week1.Rows(4).Cells(_ColumCount).Text) = "-" Then
                            Else
                                dg_Knt_Week1.Rows(4).Cells(_ColumCount).Value = "-"

                                Value = dg_Knt_Week1.Rows(0).Cells(3).Value
                                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                                dg_Knt_Week1.Rows(5).Cells(_ColumCount).Value = _St
                                ' _ColumCount = _ColumCount + 1
                            End If
                        End If


                        _ColumCount = _ColumCount + 1

                    Else
                        If Trim(dg_Knt_Week1.Rows(_ROWCOUNT).Cells(_ColumCount).Text) = "-" Then
                        Else
                            dg_Knt_Week1.Rows(_ROWCOUNT).Cells(_ColumCount).Value = "-"

                            Value = dg_Knt_Week1.Rows(0).Cells(_ColumCount).Value
                            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                            dg_Knt_Week1.Rows(_ROWCOUNT + 1).Cells(_ColumCount).Value = _St

                            dg_Knt_Week1.Rows(4).Cells(_ColumCount).Value = "-"

                            If IsNumeric(dg_Knt_Week1.Rows(3).Cells(_ColumCount).Value) Then
                                Value = dg_Knt_Week1.Rows(3).Cells(_ColumCount).Value
                                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                                dg_Knt_Week1.Rows(5).Cells(_ColumCount).Value = _St
                            End If
                            _ColumCount = _ColumCount + 1
                            End If
                    End If
                    '--------------------------------------------------------------------------------
                    'PLANNER CONSUME
                   

                    userDate = userDate.AddDays(+7)
                Next
                I = I + 1
            Next
            '=================================================================
            'ALLOCATED PROJECTION
            'NEED TO CHECK AGAIN
            _ColumCount = 1
            Dim _ProjectCode As Integer

            vcWhere = "P01CODE='PRN'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "P01"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                _ProjectCode = M01.Tables(0).Rows(0)("P01NO")
            End If
            If _ProjectCode > 0 Then
                _ProjectCode = _ProjectCode - 1
            End If
            I = 0
            Dim _StelingQulity As String
            Dim _QTY As Double

            _QTY = 0
            'DEVELOPED BY SURANGA WIJESINGHE
            _StelingQulity = ""
            vcWhere = "T15Sales_Order='" & strSales_Order & "' AND T15Line_Item=" & strLine_Item & ""
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "BP1"), New SqlParameter("@vcWhereClause1", vcWhere))
            For Each DTRow As DataRow In M01.Tables(0).Rows
                If I = 0 Then
                    _StelingQulity = Trim(M01.Tables(0).Rows(I)("T15Quality"))
                Else
                    _StelingQulity = "','" & Trim(M01.Tables(0).Rows(I)("T15Quality"))
                End If
                I = I + 1
            Next
            I = 0

            vcWhere = "T15Sales_Order='" & strSales_Order & "' and T15Line_Item=" & txtLine_Item.Text & ""
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "PROX"), New SqlParameter("@vcWhereClause1", vcWhere))
            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                Dim remain As Integer
                Dim noOfWeek As Integer
                Dim userDate As Date
                Dim _LastDate As Date
                Dim _TimeSpan As TimeSpan

                userDate = DateTime.Parse(M01.Tables(0).Rows(I)("T15Month") & "/1/" & M01.Tables(0).Rows(I)("T15Year"))
                ' MsgBox(WeekdayName(Weekday(userDate)))
                If WeekdayName(Weekday(userDate)) = "Sunday" Then
                    userDate = userDate.AddDays(-3)
                ElseIf WeekdayName(Weekday(userDate)) = "Monday" Then
                    userDate = userDate.AddDays(-4)
                ElseIf WeekdayName(Weekday(userDate)) = "Tuesday" Then
                    userDate = userDate.AddDays(-5)
                ElseIf WeekdayName(Weekday(userDate)) = "Thusday" Then
                    'userDate = userDate.AddDays(-1)
                ElseIf WeekdayName(Weekday(userDate)) = "Friday" Then
                    userDate = userDate.AddDays(-1)
                ElseIf WeekdayName(Weekday(userDate)) = "Saturday" Then
                    userDate = userDate.AddDays(-2)

                End If
                _LastDate = DateTime.Parse(M01.Tables(0).Rows(I)("T15Month") & "/1/" & M01.Tables(0).Rows(I)("T15Year"))
                ' MsgBox(Date.DaysInMonth(_LastDate.Year, _LastDate.Month))
                _LastDate = DateTime.Parse(M01.Tables(0).Rows(I)("T15Month") & "/" & Date.DaysInMonth(_LastDate.Year, _LastDate.Month) & "/" & M01.Tables(0).Rows(I)("T15Year"))
                If WeekdayName(Weekday(_LastDate)) = "Sunday" Then
                    _LastDate = _LastDate.AddDays(-4)
                ElseIf WeekdayName(Weekday(_LastDate)) = "Monday" Then
                    _LastDate = _LastDate.AddDays(-5)
                ElseIf WeekdayName(Weekday(_LastDate)) = "Tuesday" Then
                    _LastDate = _LastDate.AddDays(-6)
                ElseIf WeekdayName(Weekday(_LastDate)) = "Thusday" Then
                    'userDate = userDate.AddDays(-1)
                ElseIf WeekdayName(Weekday(_LastDate)) = "Friday" Then
                    _LastDate = _LastDate.AddDays(-2)
                ElseIf WeekdayName(Weekday(_LastDate)) = "Saturday" Then
                    _LastDate = _LastDate.AddDays(-3)

                End If

                _TimeSpan = _LastDate.Subtract(userDate)
                _WeekNo = _TimeSpan.Days / 7

                vcWhere = "M43Quality IN ('" & _StelingQulity & "') and M43Count_No=" & _ProjectCode & " and M43Year=" & M01.Tables(0).Rows(I)("T15Year") & " and M43Product_Month=" & M01.Tables(0).Rows(I)("T15Month") & ""
                dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TMP5"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(dsUser) Then
                    _QTY = dsUser.Tables(0).Rows(0)("QTY")
                  
                End If
               

                If _QTY > 0 Then
                    _QTY = _QTY / _WeekNo
                End If

                _ColumCount = 1
                For Z = 1 To _WeekNo
                    Dim _St As String
                    Value = _QTY
                    _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                    dg_Knt_Week1.Rows(7).Cells(_ColumCount).Value = _St
                    _ColumCount = _ColumCount + 1
                Next
                I = I + 1
            Next
            ' MsgBox(Delivary_Ref)
           

            'WEEKLY CONSUME PROJECTION
            _ColumCount = 1
            I = 0
            vcWhere = "T15Sales_Order='" & strSales_Order & "' AND T15Line_Item='" & strLine_Item & "'"
            T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TMP7"), New SqlParameter("@vcWhereClause1", vcWhere))
            For Each DTRow3 As DataRow In T01.Tables(0).Rows

                Dim remain As Integer
                Dim noOfWeek As Integer
                Dim userDate As Date
                Dim _LastDate As Date
                Dim _TimeSpan As TimeSpan

                userDate = DateTime.Parse(T01.Tables(0).Rows(I)("T15Month") & "/1/" & T01.Tables(0).Rows(I)("T15Year"))
                ' MsgBox(WeekdayName(Weekday(userDate)))
                If WeekdayName(Weekday(userDate)) = "Sunday" Then
                    userDate = userDate.AddDays(-3)
                ElseIf WeekdayName(Weekday(userDate)) = "Monday" Then
                    userDate = userDate.AddDays(-4)
                ElseIf WeekdayName(Weekday(userDate)) = "Tuesday" Then
                    userDate = userDate.AddDays(-5)
                ElseIf WeekdayName(Weekday(userDate)) = "Thusday" Then
                    'userDate = userDate.AddDays(-1)
                ElseIf WeekdayName(Weekday(userDate)) = "Friday" Then
                    userDate = userDate.AddDays(-1)
                ElseIf WeekdayName(Weekday(userDate)) = "Saturday" Then
                    userDate = userDate.AddDays(-2)

                End If
                _LastDate = DateTime.Parse(T01.Tables(0).Rows(I)("T15Month") & "/1/" & T01.Tables(0).Rows(I)("T15Year"))
                ' MsgBox(Date.DaysInMonth(_LastDate.Year, _LastDate.Month))
                _LastDate = DateTime.Parse(T01.Tables(0).Rows(I)("T15Month") & "/" & Date.DaysInMonth(_LastDate.Year, _LastDate.Month) & "/" & T01.Tables(0).Rows(I)("T15Year"))

                If WeekdayName(Weekday(_LastDate)) = "Sunday" Then
                    _LastDate = _LastDate.AddDays(-4)
                ElseIf WeekdayName(Weekday(_LastDate)) = "Monday" Then
                    _LastDate = _LastDate.AddDays(-5)
                ElseIf WeekdayName(Weekday(_LastDate)) = "Tuesday" Then
                    _LastDate = _LastDate.AddDays(-6)
                ElseIf WeekdayName(Weekday(_LastDate)) = "Thusday" Then
                    'userDate = userDate.AddDays(-1)
                ElseIf WeekdayName(Weekday(_LastDate)) = "Friday" Then
                    _LastDate = _LastDate.AddDays(-2)
                ElseIf WeekdayName(Weekday(_LastDate)) = "Saturday" Then
                    _LastDate = _LastDate.AddDays(-3)

                End If

                _TimeSpan = _LastDate.Subtract(userDate)
                _WeekNo = _TimeSpan.Days / 7

                userDate = userDate.AddDays(+7)
                _ColumCount = 1
                For Z = 1 To _WeekNo

                    Dim culture As System.Globalization.CultureInfo
                    Dim intWeek As Integer
                    Dim _StrWeek As String
                    Dim _StrWeek1 As String



                    culture = System.Globalization.CultureInfo.CurrentCulture
                    intWeek = culture.Calendar.GetWeekOfYear(userDate, System.Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday)

                    vcWhere = "T15Code='" & Trim(dg_Knt_Pojection.Rows(0).Cells(0).Text) & "' AND T15Shade='" & Trim(dg_Knt_Pojection.Rows(0).Cells(1).Text) & "' AND tmpWeek_No=" & intWeek & " AND tmpYear=" & Year(userDate) & ""
                    dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "KWAL"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(dsUser) Then
                        Dim _St As String
                        Value = dsUser.Tables(0).Rows(0)("Qty")
                        _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                        dg_Knt_Week1.Rows(8).Cells(_ColumCount).Value = _St

                        Value = dg_Knt_Week1.Rows(7).Cells(_ColumCount).Value - Value
                        _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                        dg_Knt_Week1.Rows(9).Cells(_ColumCount).Value = _St
                        _ColumCount = _ColumCount + 1
                    Else
                        dg_Knt_Week1.Rows(9).Cells(_ColumCount).Value = dg_Knt_Week1.Rows(7).Cells(_ColumCount).Value
                        _ColumCount = _ColumCount + 1
                    End If
                    userDate = userDate.AddDays(+6)
                Next
                I = I + 1
            Next

            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            con.close()
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
                con.close()

            End If
        End Try
    End Function

    Function Load_Pro_YD_Plan()
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

        Dim agroup1_1 As UltraGridGroup
        Dim agroup1_2 As UltraGridGroup
        Dim agroup1_3 As UltraGridGroup
        Dim agroup1_4 As UltraGridGroup
        Dim agroup1_5 As UltraGridGroup

        Dim agroup2_1 As UltraGridGroup
        Dim agroup2_2 As UltraGridGroup
        Dim agroup2_3 As UltraGridGroup
        Dim agroup2_4 As UltraGridGroup
        Dim agroup2_5 As UltraGridGroup

        Dim agroup3_1 As UltraGridGroup
        Dim agroup3_2 As UltraGridGroup
        Dim agroup3_3 As UltraGridGroup
        Dim agroup3_4 As UltraGridGroup
        Dim agroup3_5 As UltraGridGroup

        Dim agroup4_1 As UltraGridGroup
        Dim agroup4_2 As UltraGridGroup
        Dim agroup4_3 As UltraGridGroup
        Dim agroup4_4 As UltraGridGroup
        Dim agroup4_5 As UltraGridGroup

        Dim agroup5_1 As UltraGridGroup
        Dim agroup5_2 As UltraGridGroup
        Dim agroup5_3 As UltraGridGroup
        Dim agroup5_4 As UltraGridGroup
        Dim agroup5_5 As UltraGridGroup

        Dim _Date As Date
        Dim countdays As Integer
        Dim _WeekNo As Integer
        Dim _FromDate As Date
        Dim _ColumCount As Integer
        Dim Value As Double
        Dim Value3 As Double

        Dim _STSting As String
        Dim _StartDate As Date
        Dim _Shade As String


        Try
            'Dim agroup1 As UltraGridGroup
            'Dim agroup2 As UltraGridGroup
            'Dim agroup3 As UltraGridGroup
            'Dim agroup4 As UltraGridGroup
            'Dim agroup5 As UltraGridGroup
            '  Dim agroup6 As UltraGridGroup

            'If dg1_YDP.DisplayLayout.Bands(0).GroupHeadersVisible = True Then
            '    dg1_YDP.DisplayLayout.Bands(0).Groups.Remove("GroupH")
            'Else
            dg1_YDP.DisplayLayout.Bands(0).Groups.Clear()
            dg1_YDP.DisplayLayout.Bands(0).Columns.Dispose()
            'End If
            ' dg1_YDP.DisplayLayout.Bands(0).Groups.Remove("GroupH")
            agroup1 = dg1_YDP.DisplayLayout.Bands(0).Groups.Add("GroupH")
            'Me.UltraGrid3.DisplayLayout.Bands(0).Columns.Add("Line", "Line Item")
            'Me.UltraGrid3.DisplayLayout.Bands(0).Columns("Line").Group = agroup1
            'Me.UltraGrid3.DisplayLayout.Bands(0).Columns("Line").Width = 50

            'Me.UltraGrid3.DisplayLayout.Bands(0).Columns.Add("##", "##")
            'Me.UltraGrid3.DisplayLayout.Bands(0).Columns("##").Group = agroup1
            'Me.UltraGrid3.DisplayLayout.Bands(0).Columns("##").Width = 120
            ''  End If
            ' agroup1 = UltraGrid3.DisplayLayout.Bands(0).Groups.Remove(0)


            _Code = _Code - 1
            agroup1.Header.Caption = ""

            agroup1.Width = 110
            Dim dt As DataTable = New DataTable()
            ' dt.Columns.Add("ID", GetType(Integer))
            Dim colWork As New DataColumn("##", GetType(String))
            dt.Columns.Add(colWork)
            colWork.ReadOnly = True


            dt.Rows.Add("Confimed Projection")
            dt.Rows.Add("Consumed Projection")
            dt.Rows.Add("Balance Projection")
            dt.Rows.Add("Projection (" & strDisname & ")")
            dt.Rows.Add("Consumed Projection(" & strDisname & ")")
            dt.Rows.Add("Balance Projection(" & strDisname & ")")
            dt.Rows.Add("")
            dt.Rows.Add("Allocated Projection ")
            dt.Rows.Add("Consumed Projection")
            dt.Rows.Add("Balance Projection")



            Me.dg1_YDP.SetDataBinding(dt, Nothing)
            Me.dg1_YDP.DisplayLayout.Bands(0).Columns(0).Group = agroup1
            'Me.dg_Knt_Week.DisplayLayout.Bands(0).Columns(1).Group = agroup1
            'Me.dg_Knt_Week.DisplayLayout.Bands(0).Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            Me.dg1_YDP.DisplayLayout.Bands(0).Columns(0).Width = 180
            ' Me.dg_Knt_Week.DisplayLayout.Bands(0).Columns(1).Width = 50
            Dim _Group As String
            'agroup2.Key = ""
            'agroup3.Key = ""
            'agroup4.Key = ""
            '' agroup5.Key = ""


            I = 0
            'T01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            vcWhere = "T15Sales_Order ='" & strSales_Order & "' and T15Line_Item='" & strLine_Item & "' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "BP7"), New SqlParameter("@vcWhereClause1", vcWhere))
            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                _Group = "Group" & I + 1
                'If I = 0 Then
                '    _FromDate = M01.Tables(0).Rows(I)("M43Product_Month") & "/1/" & M01.Tables(0).Rows(I)("M43Year")
                '    _StartDate = _FromDate
                '    _FromDate = _FromDate.AddDays(-21)
                '    _WeekNo = weekNumber(_FromDate)
                'End If

                agroup2 = dg1_YDP.DisplayLayout.Bands(0).Groups.Add(_Group)

                agroup2.Header.Caption = Trim(M01.Tables(0).Rows(I)("M13Name"))


                countdays = DateTime.DaysInMonth(M01.Tables(0).Rows(I)("M43Year"), M01.Tables(0).Rows(I)("M43Product_Month"))
                countdays = countdays / 7


                If countdays = 4 Then
                    agroup2.Width = 200
                ElseIf countdays = 5 Then
                    agroup2.Width = 250
                End If

                Dim remain As Integer
                Dim noOfWeek As Integer
                Dim userDate As Date
                Dim _LastDate As Date
                Dim _TimeSpan As TimeSpan

                userDate = DateTime.Parse(M01.Tables(0).Rows(I)("M43Product_Month") & "/1/" & M01.Tables(0).Rows(I)("M43Year"))
                ' MsgBox(WeekdayName(Weekday(userDate)))
                If WeekdayName(Weekday(userDate)) = "Sunday" Then
                    userDate = userDate.AddDays(-3)
                ElseIf WeekdayName(Weekday(userDate)) = "Monday" Then
                    userDate = userDate.AddDays(-4)
                ElseIf WeekdayName(Weekday(userDate)) = "Tuesday" Then
                    userDate = userDate.AddDays(-5)
                ElseIf WeekdayName(Weekday(userDate)) = "Thusday" Then
                    'userDate = userDate.AddDays(-1)
                ElseIf WeekdayName(Weekday(userDate)) = "Friday" Then
                    userDate = userDate.AddDays(-1)
                ElseIf WeekdayName(Weekday(userDate)) = "Saturday" Then
                    userDate = userDate.AddDays(-2)

                End If
                _LastDate = DateTime.Parse(M01.Tables(0).Rows(I)("M43Product_Month") & "/1/" & M01.Tables(0).Rows(I)("M43Year"))
                ' MsgBox(Date.DaysInMonth(_LastDate.Year, _LastDate.Month))
                _LastDate = DateTime.Parse(M01.Tables(0).Rows(I)("M43Product_Month") & "/" & Date.DaysInMonth(_LastDate.Year, _LastDate.Month) & "/" & M01.Tables(0).Rows(I)("M43Year"))

                If WeekdayName(Weekday(_LastDate)) = "Sunday" Then
                    _LastDate = _LastDate.AddDays(-4)
                ElseIf WeekdayName(Weekday(_LastDate)) = "Monday" Then
                    _LastDate = _LastDate.AddDays(-5)
                ElseIf WeekdayName(Weekday(_LastDate)) = "Tuesday" Then
                    _LastDate = _LastDate.AddDays(-6)
                ElseIf WeekdayName(Weekday(_LastDate)) = "Thusday" Then
                    'userDate = userDate.AddDays(-1)
                ElseIf WeekdayName(Weekday(_LastDate)) = "Friday" Then
                    _LastDate = _LastDate.AddDays(-2)
                ElseIf WeekdayName(Weekday(_LastDate)) = "Saturday" Then
                    _LastDate = _LastDate.AddDays(-3)

                End If

                _TimeSpan = _LastDate.Subtract(userDate)
                _WeekNo = _TimeSpan.Days / 7

                userDate = userDate.AddDays(+7)
                vcWhere = "T15Sales_Order='" & Trim(cboSO.Text) & "' AND t01bulk='1st Bulk' and T15Year=" & M01.Tables(0).Rows(I)("M43Year") & " and T15Month=" & M01.Tables(0).Rows(I)("M43Product_Month") & ""
                T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "FSTB"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(T01) Then
                    userDate = userDate.AddDays(-35)
                Else
                    userDate = userDate.AddDays(-28)
                End If

                If _WeekNo = 5 Then
                    Dim culture As System.Globalization.CultureInfo
                    Dim intWeek As Integer
                    Dim _StrWeek As String
                    Dim _StrWeek1 As String

                    culture = System.Globalization.CultureInfo.CurrentCulture
                    intWeek = culture.Calendar.GetWeekOfYear(userDate, System.Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday)

                    _StrWeek = "Week " & intWeek
                    _StrWeek1 = "Week " & intWeek
                    Me.dg1_YDP.DisplayLayout.Bands(0).Columns.Add(_StrWeek1, _StrWeek)
                    Me.dg1_YDP.DisplayLayout.Bands(0).Columns(_StrWeek1).Group = agroup2
                    Me.dg1_YDP.DisplayLayout.Bands(0).Columns(_StrWeek1).Width = 60
                    Me.dg1_YDP.DisplayLayout.Bands(0).Columns(_StrWeek1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                    userDate = userDate.AddDays(+7)
                    culture = System.Globalization.CultureInfo.CurrentCulture
                    intWeek = culture.Calendar.GetWeekOfYear(userDate, System.Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday)

                    _StrWeek = "Week " & intWeek
                    _StrWeek1 = "Week " & intWeek

                    Me.dg1_YDP.DisplayLayout.Bands(0).Columns.Add(_StrWeek1, _StrWeek)
                    Me.dg1_YDP.DisplayLayout.Bands(0).Columns(_StrWeek1).Group = agroup2
                    Me.dg1_YDP.DisplayLayout.Bands(0).Columns(_StrWeek1).Width = 60
                    Me.dg1_YDP.DisplayLayout.Bands(0).Columns(_StrWeek1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                    userDate = userDate.AddDays(+7)
                    culture = System.Globalization.CultureInfo.CurrentCulture
                    intWeek = culture.Calendar.GetWeekOfYear(userDate, System.Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday)

                    _StrWeek = "Week " & intWeek
                    _StrWeek1 = "Week " & intWeek
                    Me.dg1_YDP.DisplayLayout.Bands(0).Columns.Add(_StrWeek1, _StrWeek)
                    Me.dg1_YDP.DisplayLayout.Bands(0).Columns(_StrWeek1).Group = agroup2
                    Me.dg1_YDP.DisplayLayout.Bands(0).Columns(_StrWeek1).Width = 60
                    Me.dg1_YDP.DisplayLayout.Bands(0).Columns(_StrWeek1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                    userDate = userDate.AddDays(+7)
                    culture = System.Globalization.CultureInfo.CurrentCulture
                    intWeek = culture.Calendar.GetWeekOfYear(userDate, System.Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday)

                    _StrWeek = "Week " & intWeek
                    _StrWeek1 = "Week " & intWeek
                    Me.dg1_YDP.DisplayLayout.Bands(0).Columns.Add(_StrWeek1, _StrWeek)
                    Me.dg1_YDP.DisplayLayout.Bands(0).Columns(_StrWeek1).Group = agroup2
                    Me.dg1_YDP.DisplayLayout.Bands(0).Columns(_StrWeek1).Width = 60
                    Me.dg1_YDP.DisplayLayout.Bands(0).Columns(_StrWeek1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                    userDate = userDate.AddDays(+7)
                    culture = System.Globalization.CultureInfo.CurrentCulture
                    intWeek = culture.Calendar.GetWeekOfYear(userDate, System.Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday)

                    _StrWeek = "Week " & intWeek
                    _StrWeek1 = "Week " & intWeek

                    Me.dg1_YDP.DisplayLayout.Bands(0).Columns.Add(_StrWeek1, _StrWeek)
                    Me.dg1_YDP.DisplayLayout.Bands(0).Columns(_StrWeek1).Group = agroup2
                    Me.dg1_YDP.DisplayLayout.Bands(0).Columns(_StrWeek1).Width = 60
                    Me.dg1_YDP.DisplayLayout.Bands(0).Columns(_StrWeek1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                ElseIf _WeekNo = 4 Then
                    Dim culture As System.Globalization.CultureInfo
                    Dim intWeek As Integer
                    Dim _StrWeek As String

                    culture = System.Globalization.CultureInfo.CurrentCulture
                    intWeek = culture.Calendar.GetWeekOfYear(userDate, System.Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday)

                    _StrWeek = "Week " & intWeek
                    Me.dg1_YDP.DisplayLayout.Bands(0).Columns.Add(_StrWeek, _StrWeek)
                    Me.dg1_YDP.DisplayLayout.Bands(0).Columns(_StrWeek).Group = agroup2
                    Me.dg1_YDP.DisplayLayout.Bands(0).Columns(_StrWeek).Width = 60
                    Me.dg1_YDP.DisplayLayout.Bands(0).Columns(_StrWeek).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                    userDate = userDate.AddDays(+7)
                    culture = System.Globalization.CultureInfo.CurrentCulture
                    intWeek = culture.Calendar.GetWeekOfYear(userDate, System.Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday)

                    _StrWeek = "Week " & intWeek
                    Me.dg1_YDP.DisplayLayout.Bands(0).Columns.Add(_StrWeek, _StrWeek)
                    Me.dg1_YDP.DisplayLayout.Bands(0).Columns(_StrWeek).Group = agroup2
                    Me.dg1_YDP.DisplayLayout.Bands(0).Columns(_StrWeek).Width = 60
                    Me.dg1_YDP.DisplayLayout.Bands(0).Columns(_StrWeek).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                    userDate = userDate.AddDays(+7)
                    culture = System.Globalization.CultureInfo.CurrentCulture
                    intWeek = culture.Calendar.GetWeekOfYear(userDate, System.Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday)

                    _StrWeek = "Week " & intWeek
                    Me.dg1_YDP.DisplayLayout.Bands(0).Columns.Add(_StrWeek, _StrWeek)
                    Me.dg1_YDP.DisplayLayout.Bands(0).Columns(_StrWeek).Group = agroup2
                    Me.dg1_YDP.DisplayLayout.Bands(0).Columns(_StrWeek).Width = 60
                    Me.dg1_YDP.DisplayLayout.Bands(0).Columns(_StrWeek).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                    userDate = userDate.AddDays(+7)
                    culture = System.Globalization.CultureInfo.CurrentCulture
                    intWeek = culture.Calendar.GetWeekOfYear(userDate, System.Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday)

                    _StrWeek = "Week " & intWeek
                    Me.dg1_YDP.DisplayLayout.Bands(0).Columns.Add(_StrWeek, _StrWeek)
                    Me.dg1_YDP.DisplayLayout.Bands(0).Columns(_StrWeek).Group = agroup2
                    Me.dg1_YDP.DisplayLayout.Bands(0).Columns(_StrWeek).Width = 60
                    Me.dg1_YDP.DisplayLayout.Bands(0).Columns(_StrWeek).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                End If

                
                I = I + 1
            Next


            vcWhere = "select * from P01PARAMETER where P01CODE='PRN'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, vcWhere)
            If isValidDataset(M01) Then
                _Code = M01.Tables(0).Rows(0)("P01NO")
            End If

            _Code = _Code - 1

            Dim Z As Integer
            Dim Value1 As Double
            Dim Value2 As Double
            Dim Y As Integer
            Dim Value5 As Double
            Dim Value6 As Double
            Dim Value7 As Double

            vcWhere = "T15Sales_Order ='" & strSales_Order & "' and T15Line_Item='" & strLine_Item & "' "
            dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "BP7"), New SqlParameter("@vcWhereClause1", vcWhere))
            For Each DTRow5 As DataRow In dsUser.Tables(0).Rows
                I = 0
                _ColumCount = 1
                Value2 = 0
                'vcWhere = "T15Sales_Order ='" & strSales_Order & "' and T15Line_Item='" & strLine_Item & "' and T15Quality='" & txtQuality.Text & "'"
                'M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TH3"), New SqlParameter("@vcWhereClause1", vcWhere))
                'If isValidDataset(M01) Then
                '    _Shade = Trim(M01.Tables(0).Rows(0)("T15Shade"))
                'End If
                vcWhere = "M43Product_Month=" & dsUser.Tables(0).Rows(Y)("M43Product_Month") & " and M43Year=" & dsUser.Tables(0).Rows(Y)("M43Year") & " and M43Count_No=" & _Code & " and M43Product_type='Yarn Dye'"
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "MPRO"), New SqlParameter("@vcWhereClause1", vcWhere))
                For Each DTRow3 As DataRow In M01.Tables(0).Rows


                    _StartDate = M01.Tables(0).Rows(I)("M43Product_Month") & "/1/" & M01.Tables(0).Rows(I)("M43Year")
                    countdays = DateTime.DaysInMonth(M01.Tables(0).Rows(I)("M43Year"), M01.Tables(0).Rows(I)("M43Product_Month"))
                    Value1 = 0
                    vcWhere = "M43Count_No=" & _Code & " and left(M43Quality,1)='Y'  "
                    T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "MPRZ"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(T01) Then
                        Value1 = T01.Tables(0).Rows(0)("Qty")
                    End If
                    Value = M01.Tables(0).Rows(0)("M43Sales_Volume")
                    countdays = countdays / 7

                    Value = Value / countdays
                    Value1 = Value1 / countdays


                    Value3 = 0
                    vcWhere = "M43Product_Month=" & M01.Tables(0).Rows(I)("M43Product_Month") & " and M43Year=" & M01.Tables(0).Rows(I)("M43Year") & " and M43Count_No=" & _Code & " and M43Product_type='Yarn Dye' and M43Planner='" & strDisname & "'"
                    T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "MPRO"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(T01) Then
                        Value3 = T01.Tables(0).Rows(0)("M43Sales_Volume")
                        Value3 = Value3 / countdays
                    End If



                    Value5 = 0
                    vcWhere = "M43Count_No=" & _Code & " and left(M43Quality,1)='Y'  and M43Planner='" & strDisname & "'"
                    T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "MPRZ"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(T01) Then
                        Value5 = T01.Tables(0).Rows(0)("Qty")
                    End If
                    Value5 = Value5 / countdays

                    Value7 = 0
                    vcWhere = "M43Year=" & M01.Tables(0).Rows(I)("M43Year") & " and M43Product_Month=" & M01.Tables(0).Rows(I)("M43Product_Month") & " and M43Quality='" & txtQuality.Text & "'  "
                    T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TH2"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(T01) Then
                        Value7 = T01.Tables(0).Rows(0)("Qty")
                    End If
                    Value7 = Value7 / countdays

                    _ColumCount = 1
                    If countdays = 4 Then

                        For Z = 0 To countdays - 1

                            If Z = 0 Then
                                _WeekNo = weekNumber(_FromDate)
                            End If
                            Value2 = 0
                            vcWhere = "tmpWeek_No=" & _WeekNo & " and tmpYear=" & Year(_FromDate) & "  "
                            T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TH1"), New SqlParameter("@vcWhereClause1", vcWhere))
                            If isValidDataset(T01) Then
                                Value2 = T01.Tables(0).Rows(0)("QTY")
                                'Value2 = Value2 / countdays
                                'Value2 = Value2 * 7
                            End If

                            _STSting = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                            _STSting = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                            dg1_YDP.Rows(0).Cells(_ColumCount).Value = CInt(Value1)

                            _STSting = (Value2.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                            _STSting = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value2))

                            dg1_YDP.Rows(1).Cells(_ColumCount).Value = CInt(Value2)

                            Value2 = Value - Value2
                            _STSting = (Value2.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                            _STSting = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value2))

                            dg1_YDP.Rows(2).Cells(_ColumCount).Value = CInt(Value2)


                            '_STSting = (Value3.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                            '_STSting = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value3))

                            'dg1_YDP.Rows(3).Cells(_ColumCount).Value = _STSting
                            'consium projection
                            'Value1 = 0
                            'vcWhere = "T15Year=" & M01.Tables(0).Rows(I)("M43Year") & " and T15Month=" & M01.Tables(0).Rows(I)("M43Product_Month") & " and left(T15Quality,1)='Y' and T15Planner='" & strDisname & "'"
                            'T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "MPR1"), New SqlParameter("@vcWhereClause1", vcWhere))
                            'If isValidDataset(T01) Then
                            '    Value1 = T01.Tables(0).Rows(0)("Qty")
                            'End If
                            'Value1 = Value1 / countdays

                            _STSting = (Value3.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                            _STSting = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value3))
                            dg1_YDP.Rows(3).Cells(_ColumCount).Value = CInt(Value5)

                            Value2 = 0
                            vcWhere = "tmpWeek_No=" & _WeekNo & " and tmpYear=" & Year(_FromDate) & " and tmpPlanner='" & strDisname & "' "
                            T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TH1"), New SqlParameter("@vcWhereClause1", vcWhere))
                            If isValidDataset(T01) Then
                                Value2 = T01.Tables(0).Rows(0)("QTY")
                                'Value2 = Value2 / countdays
                                'Value2 = Value2 * 7
                            End If
                            _STSting = (Value2.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                            _STSting = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value2))
                            dg1_YDP.Rows(4).Cells(_ColumCount).Value = CInt(Value2)

                            Value6 = Value3 - Value2
                            _STSting = (Value6.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                            _STSting = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value6))
                            dg1_YDP.Rows(5).Cells(_ColumCount).Value = CInt(Value6)
                            dg1_YDP.Rows(5).Cells(_ColumCount).Appearance.BackColor = Color.Red
                            dg1_YDP.Rows(5).Cells(0).Appearance.BackColor = Color.Red

                            dg1_YDP.Rows(2).Cells(_ColumCount).Appearance.BackColor = Color.Red
                            dg1_YDP.Rows(2).Cells(0).Appearance.BackColor = Color.Red

                            dg1_YDP.Rows(7).Cells(_ColumCount).Value = CInt(Value7)
                            dg1_YDP.Rows(9).Cells(_ColumCount).Value = CInt(Value7)
                            dg1_YDP.Rows(8).Cells(_ColumCount).Value = "0"
                            _WeekNo = _WeekNo + 1
                            _ColumCount = _ColumCount + 1
                        Next

                    ElseIf countdays = 5 Then

                        For Z = 0 To countdays

                            If Z = 0 Then
                                _WeekNo = weekNumber(_FromDate)
                            End If
                            Value2 = 0
                            vcWhere = "tmpWeek_No=" & _WeekNo & " and tmpYear=" & Year(_FromDate) & "  and tmpPlanner='" & strDisname & "' "
                            T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TH1"), New SqlParameter("@vcWhereClause1", vcWhere))
                            If isValidDataset(T01) Then
                                Value2 = T01.Tables(0).Rows(0)("QTY")
                                'Value2 = Value2 / countdays
                                'Value2 = Value2 * 7
                            End If

                            _STSting = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                            _STSting = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                            dg1_YDP.Rows(0).Cells(_ColumCount).Value = CInt(Value)

                            _STSting = (Value2.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                            _STSting = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value2))

                            dg1_YDP.Rows(1).Cells(_ColumCount).Value = CInt(Value2)

                            Value2 = Value - Value2
                            _STSting = (Value2.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                            _STSting = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value2))

                            dg1_YDP.Rows(2).Cells(_ColumCount).Value = CInt(Value2)


                            '_STSting = (Value3.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                            '_STSting = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value3))

                            'dg1_YDP.Rows(3).Cells(_ColumCount).Value = _STSting
                            'consium projection
                            'Value1 = 0
                            'vcWhere = "T15Year=" & M01.Tables(0).Rows(I)("M43Year") & " and T15Month=" & M01.Tables(0).Rows(I)("M43Product_Month") & " and left(T15Quality,1)='Y' and T15Planner='" & strDisname & "'"
                            'T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "MPR1"), New SqlParameter("@vcWhereClause1", vcWhere))
                            'If isValidDataset(T01) Then
                            '    Value1 = T01.Tables(0).Rows(0)("Qty")
                            'End If
                            'Value1 = Value1 / countdays

                            _STSting = (Value5.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                            _STSting = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value5))
                            dg1_YDP.Rows(3).Cells(_ColumCount).Value = CInt(Value3)

                            Value2 = 0
                            vcWhere = "tmpWeek_No=" & _WeekNo & " and tmpYear=" & Year(_FromDate) & " and tmpPlanner='" & strDisname & "' "
                            T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TH1"), New SqlParameter("@vcWhereClause1", vcWhere))
                            If isValidDataset(T01) Then
                                Value2 = T01.Tables(0).Rows(0)("QTY")
                                'Value2 = Value2 / countdays
                                'Value2 = Value2 * 7
                            End If
                            _STSting = (Value2.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                            _STSting = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value2))
                            dg1_YDP.Rows(4).Cells(_ColumCount).Value = CInt(Value2)

                            Value6 = Value3 - Value2
                            _STSting = (Value6.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                            _STSting = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value6))
                            dg1_YDP.Rows(5).Cells(_ColumCount).Value = CInt(Value6)
                            dg1_YDP.Rows(5).Cells(_ColumCount).Appearance.BackColor = Color.Red
                            dg1_YDP.Rows(5).Cells(_ColumCount + 1).Appearance.BackColor = Color.Red
                            dg1_YDP.Rows(5).Cells(0).Appearance.BackColor = Color.Red

                            dg1_YDP.Rows(2).Cells(_ColumCount).Appearance.BackColor = Color.Red
                            dg1_YDP.Rows(2).Cells(0).Appearance.BackColor = Color.Red

                            dg1_YDP.Rows(7).Cells(_ColumCount).Value = CInt(Value7)
                            dg1_YDP.Rows(8).Cells(_ColumCount).Value = CInt(Value7)

                            _WeekNo = _WeekNo + 1
                            _ColumCount = _ColumCount + 1
                        Next
                    End If
                    I = I + 1
                Next
                Y = Y + 1
            Next
            con.close()
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
                con.close()

            End If
        End Try
    End Function

    Function Quality_Group()
        Dim vcWhere As String
        Dim M01 As DataSet
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Try
            vcWhere = "M22Quality='" & Trim(txtQuality.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetQuality_Group", New SqlParameter("@cQryType", "SNL"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                txtQuality_Group.Text = "Solid Non Lycra"
                con.close()
                Exit Function
            End If
            '-----------------------------------------------------------------------
            'SOLID LYCRA HEVY
            vcWhere = "M22Quality='" & Trim(txtQuality.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetQuality_Group", New SqlParameter("@cQryType", "SLH"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                txtQuality_Group.Text = "Solid Lycra Heavy"
                con.close()
                Exit Function
            End If
            '-----------------------------------------------------------------------
            'SOLID LYCRA SLACK
            vcWhere = "M22Quality='" & Trim(txtQuality.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetQuality_Group", New SqlParameter("@cQryType", "SLS"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                txtQuality_Group.Text = "Solid Lycra Slack"
                con.close()
                Exit Function
            End If

            'SOLID NON LYCRA HAVY
            vcWhere = "M22Quality='" & Trim(txtQuality.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetQuality_Group", New SqlParameter("@cQryType", "SNLH"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                txtQuality_Group.Text = "Solid Non Lycra Heavy"
                con.close()
                Exit Function
            End If
            '----------------------------------------------------------------------
            'MARL NON LYCRA
            vcWhere = "M22Quality='" & Trim(txtQuality.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetQuality_Group", New SqlParameter("@cQryType", "MNL"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                txtQuality_Group.Text = "Marl Non Lycra"
                con.close()
                Exit Function
            End If
            '----------------------------------------------------------------------
            'MARL LYCRA
            vcWhere = "M22Quality='" & Trim(txtQuality.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetQuality_Group", New SqlParameter("@cQryType", "MLC"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                txtQuality_Group.Text = "Marl Lycra"
                con.close()
                Exit Function
            End If
            '----------------------------------------------------------------------
            'Dye Yarn Lycra 
            vcWhere = "M22Quality='" & Trim(txtQuality.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetQuality_Group", New SqlParameter("@cQryType", "DYL"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                txtQuality_Group.Text = "Dye Yarn Lycra"
                con.close()
                Exit Function
            End If
            '----------------------------------------------------------------------
            'Dye Yarn Non Lycra 
            vcWhere = "M22Quality='" & Trim(txtQuality.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetQuality_Group", New SqlParameter("@cQryType", "DYNL"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                txtQuality_Group.Text = "Dye Yarn Non Lycra"
                con.close()
                Exit Function
            End If

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Function Calculation_Balance_YB()
        Dim I As Integer
        Dim Value As Double
        Dim _Vstring As String
        Dim _RowIndex As Integer
        Dim _Value1 As Integer
        Dim _Value2 As Integer
        Dim _Value3 As Integer


        Try
            '  MsgBox(UltraGrid1.Selected.Rows.Item(0).ToString)
            I = 0
            Value = 0
            _Value1 = 0
            _Value2 = 0
            _Value3 = 0

            For Each uRow As UltraGridRow In dg1_YB.Rows

                With dg1_YB
                    If IsNumeric(.Rows(I).Cells(6).Text) Then

                        If CDbl((.Rows(I).Cells(6).Text)) <= CDbl((.Rows(I).Cells(5).Text)) Then
                            If Microsoft.VisualBasic.Left((.Rows(I).Cells(1).Text), 4) = Microsoft.VisualBasic.Left(txtYarn1.Text, 4) Then
                                If IsNumeric(.Rows(I).Cells(6).Text) Then
                                    _Value1 = _Value1 + .Rows(I).Cells(6).Text
                                End If
                            ElseIf Microsoft.VisualBasic.Left((.Rows(I).Cells(1).Text), 4) = Microsoft.VisualBasic.Left(txtYarn2.Text, 4) Then
                                If IsNumeric(.Rows(I).Cells(6).Text) Then
                                    _Value2 = _Value2 + .Rows(I).Cells(6).Text
                                End If
                            ElseIf Microsoft.VisualBasic.Left((.Rows(I).Cells(1).Text), 4) = Microsoft.VisualBasic.Left(txtYarn3.Text, 4) Then
                                If IsNumeric(.Rows(I).Cells(6).Text) Then
                                    _Value3 = _Value3 + .Rows(I).Cells(6).Text
                                End If
                            End If
                            Value = Value + CDbl((.Rows(I).Cells(6).Text))
                            .Rows(I).Cells(0).Appearance.BackColor = Color.White
                            .Rows(I).Cells(1).Appearance.BackColor = Color.White
                            .Rows(I).Cells(2).Appearance.BackColor = Color.White
                            .Rows(I).Cells(3).Appearance.BackColor = Color.White
                            .Rows(I).Cells(4).Appearance.BackColor = Color.White
                            .Rows(I).Cells(5).Appearance.BackColor = Color.White
                            .Rows(I).Cells(6).Appearance.BackColor = Color.White

                        Else
                            MsgBox("Qty grater than to stock", MsgBoxStyle.Information, "Information ....")
                            .Rows(I).Cells(0).Appearance.BackColor = Color.Red
                            .Rows(I).Cells(1).Appearance.BackColor = Color.Red
                            .Rows(I).Cells(2).Appearance.BackColor = Color.Red
                            .Rows(I).Cells(3).Appearance.BackColor = Color.Red
                            .Rows(I).Cells(4).Appearance.BackColor = Color.Red
                            .Rows(I).Cells(5).Appearance.BackColor = Color.Red
                            .Rows(I).Cells(6).Appearance.BackColor = Color.Red
                            .Rows(I).Selected = True
                            Exit For
                        End If
                        '_Vstring = Value
                        '_Vstring = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        '_Vstring = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                        'UltraGrid1.Rows(I).Cells(6).Value = _Vstring
                    End If
                End With
                I = I + 1
            Next
            Value = CDbl(txtReq_Grige_YB.Text) - Value
            lblBalance_YB.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            lblBalance_YB.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

            'pbCount1.Value = _Value1
            'pbCount2.Value = _Value2
            'pbCount3.Value = _Value3


            If CDbl(lblBalance_YB.Text) < 0 Then
                MsgBox("Qty grater than to stock", MsgBoxStyle.Information, "Information ....")
            End If

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                'Con.close()
            End If
        End Try
    End Function

    Function Load_Gridewith_Data_YB()
        Dim i As Integer
        Dim vcWhere As String
        Dim M01 As DataSet
        Dim M02 As DataSet
        Dim Value As Double
        Dim _VString As String
        Dim Diff As TimeSpan
        Dim _To As Date
        'Dim Value As Double
        Dim T01 As DataSet
        Dim tmpQty As Double
        Dim _Stock As String
        Dim _Quality As String

        Try
            Dim con = New SqlConnection()
            con = DBEngin.GetConnection(True)


            Dim Z As Integer
            Z = 0
            i = 0
            If Microsoft.VisualBasic.Left(Trim(txtQuality.Text), 1) = "Q" Then
                _Quality = Microsoft.VisualBasic.Right(Trim(txtQuality.Text), Microsoft.VisualBasic.Len(Trim(txtQuality.Text)) - 1)
            Else
                _Quality = Trim(txtQuality.Text)
            End If
            vcWhere = "M22Quality='" & _Quality & "' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TEC"), New SqlParameter("@vcWhereClause1", vcWhere))
            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                Dim _shade As String
                Z = 0
                If txtShade.Text = "L" Or txtShade.Text = "WS" Then
                    _shade = "('Light','SP')"
                ElseIf txtShade.Text = "D" Then
                    _shade = "('Dark','SP')"
                ElseIf txtShade.Text = "M" Then
                    _shade = "('Marl','SP')"

                End If
                If txtShade.Text <> "" Then
                    If i = 0 Then
                        If txtCom1.Text <> "" Then
                        Else
                            txtCom1.Text = "0"
                        End If
                        If txtCom1.Text > 15 Then
                            vcWhere = "dis='" & Microsoft.VisualBasic.Left(Trim(M01.Tables(0).Rows(i)("M22Yarn")), 4) & "'  and M34Shade in " & _shade & " and M33Yarn_Location in ('2020','2005','2116','2009','2110')"
                        Else
                            vcWhere = "dis='" & Microsoft.VisualBasic.Left(Trim(M01.Tables(0).Rows(i)("M22Yarn")), 4) & "' and M33Yarn_Location in ('2020','2005','2116','2009','2110')"
                        End If
                        ' End If
                    ElseIf i = 1 Then
                        If txtCom2.Text <> "" Then
                        Else
                            txtCom2.Text = "0"
                        End If
                        If txtCom2.Text > 15 Then
                            vcWhere = "dis='" & Microsoft.VisualBasic.Left(Trim(M01.Tables(0).Rows(i)("M22Yarn")), 4) & "'  and M34Shade in " & _shade & " and M33Yarn_Location in ('2020','2005','2116','2009','2110')"
                        Else
                            vcWhere = "dis='" & Microsoft.VisualBasic.Left(Trim(M01.Tables(0).Rows(i)("M22Yarn")), 4) & "' and M33Yarn_Location in ('2020','2005','2116','2009','2110')"
                        End If
                    ElseIf i = 2 Then
                        If txtCom3.Text <> "" Then
                        Else
                            txtCom3.Text = "0"
                        End If
                        If txtCom3.Text > 15 Then
                            vcWhere = "dis='" & Microsoft.VisualBasic.Left(Trim(M01.Tables(0).Rows(i)("M22Yarn")), 4) & "'  and M34Shade in " & _shade & " and M33Yarn_Location in ('2020','2005','2116','2009','2110')"
                        Else
                            vcWhere = "dis='" & Microsoft.VisualBasic.Left(Trim(M01.Tables(0).Rows(i)("M22Yarn")), 4) & "' and M33Yarn_Location in ('2020','2005','2116','2009','2110')"
                        End If
                    End If

                Else
                    vcWhere = "dis='" & Microsoft.VisualBasic.Left(Trim(M01.Tables(0).Rows(i)("M22Yarn")), 4) & "' and M33Yarn_Location in ('2020','2005','2116','2009','2110')"
                End If
                M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetYarn_Booking", New SqlParameter("@cQryType", "CLS1"), New SqlParameter("@vcWhereClause1", vcWhere))
                For Each DTRow4 As DataRow In M02.Tables(0).Rows
                    tmpQty = 0
                    Dim newRow As DataRow = c_dataCustomer1_YB.NewRow

                    newRow("10Class") = M02.Tables(0).Rows(Z)("M3310Class")
                    newRow("Description") = M02.Tables(0).Rows(Z)("M33Description")
                    newRow("Stock Code") = M02.Tables(0).Rows(Z)("M33Stock_Code")

                    newRow("Location") = M02.Tables(0).Rows(Z)("M33Yarn_Location")
                    _To = Month(M02.Tables(0).Rows(Z)("M33Date")) & "/" & Microsoft.VisualBasic.Day(M02.Tables(0).Rows(Z)("M33Date")) & "/" & Year(M02.Tables(0).Rows(Z)("M33Date"))
                    Diff = Today.Subtract(_To)
                    newRow("Yarn Aging") = Diff.Days & " days"
                    'If Diff.Days < 30 Then
                    '    newRow("Age") = "Below 1 Month"
                    'ElseIf Diff.Days >= 30 And Diff.Days < 60 Then
                    '    newRow("Age") = "Below 2 Month"
                    'ElseIf Diff.Days >= 60 And Diff.Days < 90 Then
                    '    newRow("Age") = "Below 3 Month"
                    'Else
                    '    newRow("Age") = "above 3 Month"
                    'End If
                    vcWhere = "tmpSO='" & strSales_Order & "' and tmpLine_Item=" & strLine_Item & " and tmp10Class='" & M02.Tables(0).Rows(Z)("M3310Class") & "' and tmpStock_Code='" & M02.Tables(0).Rows(Z)("M33Stock_Code") & "' and tmpLocation='" & M02.Tables(0).Rows(Z)("M33Yarn_Location") & "'"
                    T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TYBK"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(T01) Then

                        Value = T01.Tables(0).Rows(0)("tmpQty")
                        _VString = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        _VString = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                        newRow("Qty use for Order") = _VString

                    Else
                        vcWhere = "tmp10Class='" & M02.Tables(0).Rows(Z)("M3310Class") & "' and tmpStock_Code='" & M02.Tables(0).Rows(Z)("M33Stock_Code") & "' and tmpLocation='" & M02.Tables(0).Rows(Z)("M33Yarn_Location") & "'"
                        T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TYBC"), New SqlParameter("@vcWhereClause1", vcWhere))
                        If isValidDataset(T01) Then
                            tmpQty = T01.Tables(0).Rows(0)("Qty")
                        End If
                    End If
                    Value = M02.Tables(0).Rows(Z)("M33Qty") + tmpQty
                    _VString = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    _VString = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    newRow("Available Qty(Kg)") = _VString
                    newRow("##") = False
                    c_dataCustomer1_YB.Rows.Add(newRow)

                    Z = Z + 1
                Next
                Dim newRow3 As DataRow = c_dataCustomer1_YB.NewRow
                c_dataCustomer1_YB.Rows.Add(newRow3)
                i = i + 1
            Next

            'PO BASE YARN
            Dim newRow1 As DataRow = c_dataCustomer1_YB.NewRow
            c_dataCustomer1_YB.Rows.Add(newRow1)

            Z = 0
            i = 0
            vcWhere = "M22Quality='" & Trim(txtQuality.Text) & "' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TEC"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
            Else
                vcWhere = "M22Quality='" & Trim(txtCommon.Text) & "' "
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TEC"), New SqlParameter("@vcWhereClause1", vcWhere))
            End If
            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                Dim _shade As String
                Z = 0
                If txtShade.Text = "L" Or txtShade.Text = "WS" Then
                    _shade = "('Light','SP')"
                ElseIf txtShade.Text = "D" Then
                    _shade = "('Dark','SP')"
                ElseIf txtShade.Text = "M" Then
                    _shade = "('Marl','SP')"

                End If
                If txtShade.Text <> "" Then
                    If i = 0 Then
                        If txtCom1.Text <> "" Then
                        Else
                            txtCom1.Text = "0"
                        End If
                        If txtCom1.Text > 15 Then
                            vcWhere = "left(M47Description,4)='" & Microsoft.VisualBasic.Left(Trim(M01.Tables(0).Rows(i)("M22Yarn")), 4) & "'  "
                        Else
                            vcWhere = "left(M47Description,4)='" & Microsoft.VisualBasic.Left(Trim(M01.Tables(0).Rows(i)("M22Yarn")), 4) & "' "
                        End If
                    ElseIf i = 1 Then
                        If txtCom2.Text <> "" Then
                        Else
                            txtCom2.Text = "0"
                        End If
                        If txtCom2.Text > 15 Then
                            vcWhere = "left(M47Description,4)='" & Microsoft.VisualBasic.Left(Trim(M01.Tables(0).Rows(i)("M22Yarn")), 4) & "'   "
                        Else
                            vcWhere = "left(M47Description,4)='" & Microsoft.VisualBasic.Left(Trim(M01.Tables(0).Rows(i)("M22Yarn")), 4) & "'"
                        End If
                    ElseIf i = 2 Then
                        If txtCom3.Text <> "" Then
                        Else
                            txtCom3.Text = "0"
                        End If
                        If txtCom3.Text > 15 Then
                            vcWhere = "left(M47Description,4)='" & Microsoft.VisualBasic.Left(Trim(M01.Tables(0).Rows(i)("M22Yarn")), 4) & "' "
                        Else
                            vcWhere = "left(M47Description,4)='" & Microsoft.VisualBasic.Left(Trim(M01.Tables(0).Rows(i)("M22Yarn")), 4) & "' "
                        End If
                    End If

                Else
                    vcWhere = "LEFT(M47Description,4)='" & Microsoft.VisualBasic.Left(Trim(M01.Tables(0).Rows(i)("M22Yarn")), 4) & "'"
                End If
                M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "YPO"), New SqlParameter("@vcWhereClause1", vcWhere))
                For Each DTRow4 As DataRow In M02.Tables(0).Rows
                    tmpQty = 0
                    Dim newRow As DataRow = c_dataCustomer1_YB.NewRow

                    If Today > M02.Tables(0).Rows(Z)("M47Del_Date") Then
                        newRow("10Class") = M02.Tables(0).Rows(Z)("M47Material")
                        newRow("Description") = M02.Tables(0).Rows(Z)("M47Description")
                        newRow("Stock Code") = M02.Tables(0).Rows(Z)("M47PO_Order") & "-" & M02.Tables(0).Rows(Z)("M47Item")
                        _Stock = M02.Tables(0).Rows(Z)("M47PO_Order") & "-" & M02.Tables(0).Rows(Z)("M47Item")
                        newRow("Location") = "PO"
                        _To = Month(M02.Tables(0).Rows(Z)("M47Del_Date")) & "/" & Microsoft.VisualBasic.Day(M02.Tables(0).Rows(Z)("M47Del_Date")) & "/" & Year(M02.Tables(0).Rows(Z)("M47Del_Date"))
                        Diff = Today.Subtract(_To)
                        'newRow("Yarn Aging") = Diff.Days & " days"
                        'If Diff.Days < 30 Then
                        newRow("Yarn Aging") = _To
                        'ElseIf Diff.Days >= 30 And Diff.Days < 60 Then
                        '    newRow("Age") = "Below 2 Month"
                        'ElseIf Diff.Days >= 60 And Diff.Days < 90 Then
                        '    newRow("Age") = "Below 3 Month"
                        'Else
                        '    newRow("Age") = "above 3 Month"
                        'End If
                        vcWhere = "tmpSO='" & strSales_Order & "' and tmpLine_Item=" & strLine_Item & " and tmp10Class='" & M02.Tables(0).Rows(Z)("M47Material") & "' and tmpStock_Code='" & _Stock & "' and tmpLocation='PO'"
                        T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TYBK"), New SqlParameter("@vcWhereClause1", vcWhere))
                        If isValidDataset(T01) Then

                            Value = T01.Tables(0).Rows(0)("tmpQty")
                            _VString = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                            _VString = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                            newRow("Qty use for Order") = _VString

                        Else
                            vcWhere = "tmp10Class='" & M02.Tables(0).Rows(Z)("M47Material") & "' and tmpStock_Code='" & _Stock & "' and tmpLocation='PO' and tmpTime>'" & M02.Tables(0).Rows(Z)("M47Time") & "'"
                            T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TYBC"), New SqlParameter("@vcWhereClause1", vcWhere))
                            If isValidDataset(T01) Then
                                tmpQty = T01.Tables(0).Rows(0)("Qty")
                            End If
                        End If
                        Value = M02.Tables(0).Rows(Z)("M47Open_Qty") - tmpQty
                        _VString = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        _VString = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                        newRow("Available Qty(Kg)") = _VString
                        newRow("##") = False
                        c_dataCustomer1_YB.Rows.Add(newRow)
                    End If
                    Z = Z + 1
                Next

                Dim newRow22 As DataRow = c_dataCustomer1_YB.NewRow
                c_dataCustomer1_YB.Rows.Add(newRow22)
                i = i + 1
            Next
            con.close()
            Dim X As Integer
            X = 0
            For Each uRow As UltraGridRow In dg2_YB.Rows
                i = 0
                'If lblBalance.Text = "0.00" Then
                '    Exit Function
                'End If
                For Each uRow1 As UltraGridRow In dg1_YB.Rows
                    Dim _Qty As Double
                    _Qty = 0
                    With dg1_YB
                        If Trim(.Rows(i).Cells(2).Value) = Trim(dg2_YB.Rows(X).Cells(0).Value) Then
                            ' _Qty = .Rows(i).Cells(4).Value
                           
                            ' lblBalance.Text = CDbl(lblBalance.Text) - _Qty
                            .Rows(i).Cells(0).Appearance.BackColor = Color.Green
                            .Rows(i).Cells(1).Appearance.BackColor = Color.Green
                            .Rows(i).Cells(2).Appearance.BackColor = Color.Green
                            .Rows(i).Cells(3).Appearance.BackColor = Color.Green
                            .Rows(i).Cells(4).Appearance.BackColor = Color.Green
                            .Rows(i).Cells(5).Appearance.BackColor = Color.Green
                            .Rows(i).Cells(6).Appearance.BackColor = Color.Green
                          
                       
                        End If
                      


                    End With
                    i = i + 1
                Next
                X = X + 1
            Next
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                '  MessageBox.Show(returnMessage.Message)
                'Con.close()
            End If
        End Try
    End Function

    Function Load_Gridewith_DataSerch_YB()
        Dim i As Integer
        Dim vcWhere As String
        Dim M01 As DataSet
        Dim M02 As DataSet
        Dim Value As Double
        Dim _VString As String
        Dim Diff As TimeSpan
        Dim _To As Date
        'Dim Value As Double

        Try
            Dim con = New SqlConnection()
            con = DBEngin.GetConnection(True)


            Dim Z As Integer
            Z = 0
            i = 0
            vcWhere = "M22Quality = '" & Trim(txtQuality.Text) & "' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TEC"), New SqlParameter("@vcWhereClause1", vcWhere))
            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                Dim _shade As String
                Z = 0
                If txtShade.Text = "L" Or txtShade.Text = "WS" Then
                    _shade = "('Light','SP')"
                ElseIf txtShade.Text = "D" Then
                    _shade = "('Dark','SP')"
                ElseIf txtShade.Text = "M" Then
                    _shade = "('Marl','SP')"

                End If
                If txtShade.Text <> "" Then
                    If i = 0 Then
                        If txtCom1.Text > 15 Then
                            vcWhere = "dis='" & Microsoft.VisualBasic.Left(Trim(M01.Tables(0).Rows(i)("M22Yarn")), 4) & "'  and M34Shade in " & _shade & " and M33Yarn_Location in ('2020','2005','2116','2009','2110')  and M33Description like '%" & txtDis.Text & "%'"
                        Else
                            vcWhere = "dis='" & Microsoft.VisualBasic.Left(Trim(M01.Tables(0).Rows(i)("M22Yarn")), 4) & "' and M33Yarn_Location in ('2020','2005','2116','2009','2110')  and M33Description like '%" & txtDis.Text & "%'"
                        End If
                    ElseIf i = 1 Then
                        If txtCom2.Text > 15 Then
                            vcWhere = "dis='" & Microsoft.VisualBasic.Left(Trim(M01.Tables(0).Rows(i)("M22Yarn")), 4) & "'  and M34Shade in " & _shade & " and M33Yarn_Location in ('2020','2005','2116','2009','2110')  and M33Description like '%" & txtDis.Text & "%'"
                        Else
                            vcWhere = "dis='" & Microsoft.VisualBasic.Left(Trim(M01.Tables(0).Rows(i)("M22Yarn")), 4) & "' and M33Yarn_Location in ('2020','2005','2116','2009','2110')  and M33Description like '%" & txtDis.Text & "%'"
                        End If
                    ElseIf i = 2 Then
                        If txtCom3.Text > 15 Then
                            vcWhere = "dis='" & Microsoft.VisualBasic.Left(Trim(M01.Tables(0).Rows(i)("M22Yarn")), 4) & "'  and M34Shade in " & _shade & " and M33Yarn_Location in ('2020','2005','2116','2009','2110')  and M33Description like '%" & txtDis.Text & "%'"
                        Else
                            vcWhere = "dis='" & Microsoft.VisualBasic.Left(Trim(M01.Tables(0).Rows(i)("M22Yarn")), 4) & "' and M33Yarn_Location in ('2020','2005','2116','2009','2110')  and M33Description like '%" & txtDis.Text & "%'"
                        End If
                    End If

                Else
                    vcWhere = "dis like '" & Microsoft.VisualBasic.Left(Trim(M01.Tables(0).Rows(i)("M22Yarn")), 4) & "' and M33Yarn_Location in ('2020','2005','2116','2009','2110') and M33Description like '%" & txtDis.Text & "%'"
                End If
                M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetYarn_Booking", New SqlParameter("@cQryType", "CLS1"), New SqlParameter("@vcWhereClause1", vcWhere))
                For Each DTRow4 As DataRow In M02.Tables(0).Rows
                    Dim newRow As DataRow = c_dataCustomer1_YB.NewRow

                    newRow("10Class") = M02.Tables(0).Rows(Z)("M3310Class")
                    newRow("Description") = M02.Tables(0).Rows(Z)("M33Description")
                    newRow("Stock Code") = M02.Tables(0).Rows(Z)("M33Stock_Code")

                    newRow("Location") = M02.Tables(0).Rows(Z)("M33Yarn_Location")
                    _To = Month(M02.Tables(0).Rows(Z)("M33Date")) & "/" & Microsoft.VisualBasic.Day(M02.Tables(0).Rows(Z)("M33Date")) & "/" & Year(M02.Tables(0).Rows(Z)("M33Date"))
                    Diff = Today.Subtract(_To)
                    newRow("Yarn Aging") = Diff.Days & " days"
                    'If Diff.Days < 30 Then
                    '    newRow("Age") = "Below 1 Month"
                    'ElseIf Diff.Days >= 30 And Diff.Days < 60 Then
                    '    newRow("Age") = "Below 2 Month"
                    'ElseIf Diff.Days >= 60 And Diff.Days < 90 Then
                    '    newRow("Age") = "Below 3 Month"
                    'Else
                    '    newRow("Age") = "above 3 Month"
                    'End If
                    Value = M02.Tables(0).Rows(Z)("M33Qty")
                    _VString = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    _VString = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    newRow("Available Qty(Kg)") = _VString
                    newRow("##") = False
                    c_dataCustomer1_YB.Rows.Add(newRow)

                    Z = Z + 1
                Next
                Dim newRow1 As DataRow = c_dataCustomer1_YB.NewRow
                c_dataCustomer1_YB.Rows.Add(newRow1)
                i = i + 1
            Next
            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                'Con.close()
            End If
        End Try
    End Function

    Function Load_Yarn_Booking()
        On Error Resume Next
        Dim Value As Double
        Value = lblBalance.Text

        txtReq_Grige_YB.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
        txtReq_Grige_YB.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
        txtK_Wastage.Text = "2"

        Call Search_Tec_Spec()

        Dim TestString As String = Trim(txtGauge.Text)
        Dim TestArray() As String = Split(TestString)

        ' TestArray holds {"apple", "", "", "", "pear", "banana", "", ""} 
        Dim LastNonEmpty As Integer = -1
        For z1 As Integer = 0 To TestArray.Length - 1
            If TestArray(z1) <> "" Then
                LastNonEmpty += 1
                TestArray(LastNonEmpty) = TestArray(z1)
                ' If z1 = 2 Then
                ''_Quality = TestArray(LastNonEmpty)
                ''Exit For
                'End If
            End If
        Next
        strGuarge = Microsoft.VisualBasic.Left(TestArray(0), 4) & "-" & TestArray(3)
        ' frmYarn_Booking.Show()
    End Function
 
    Function Load_Grid_SockCode_YB()
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim vcWhere As String
        Dim M01 As DataSet
        Dim i As Integer
        Dim Value As Double
        Dim _VString As String

        Try
            i = 0
            vcWhere = "M42Rcode='" & Trim(txtRcode.Text) & "'  "
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetUse_Griege_Qty", New SqlParameter("@cQryType", "STC"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
            Else
                vcWhere = "m42Quality='" & Trim(txtCommon.Text) & "'  "
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetUse_Griege_Qty", New SqlParameter("@cQryType", "STC"), New SqlParameter("@vcWhereClause1", vcWhere))
            End If
           
            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer2_YB.NewRow

                newRow("Stock Code") = M01.Tables(0).Rows(i)("M42Stock_Code")
                newRow("Date") = Month(M01.Tables(0).Rows(i)("M24Customer")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("M24Customer")) & "/" & Year(M01.Tables(0).Rows(i)("M24Customer"))
                newRow("Week No") = M01.Tables(0).Rows(i)("M24Week")
                newRow("Year") = M01.Tables(0).Rows(i)("M24Year")
                Value = M01.Tables(0).Rows(i)("Qty")
                _VString = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _VString = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Used Qty(Kg)") = _VString
                c_dataCustomer2_YB.Rows.Add(newRow)
                i = i + 1
            Next
            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try

    End Function

    Private Sub cmdDye_Yarn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDye_Yarn.Click
        Dim i As Integer
        If OPR12.Visible = True Then
            OPR12.Visible = False
            UltraGroupBox36.Visible = True
        Else

            i = 0
            For Each uRow As UltraGridRow In dg1.Rows
                If Trim(dg1.Rows(i).Cells(0).Text) <> "" Then
                    If Trim(dg1.Rows(i).Cells(9).Text) <> "" Then
                        If IsNumeric(Trim(dg1.Rows(i).Cells(9).Text)) Then

                        Else
                            Dim windowInfo As New Infragistics.Win.Misc.UltraDesktopAlertShowWindowInfo
                            Dim strFileName As String
                            windowInfo.Caption = "Please check the yarn qty"
                            windowInfo.FooterText = "Technova"
                            strFileName = ConfigurationManager.AppSettings("SoundPath") + "\REMINDER.wav"
                            windowInfo.Sound = strFileName
                            UltraDesktopAlert1.Show(windowInfo)
                            Exit Sub
                        End If
                    Else
                        MsgBox("Please Request the Yarn", MsgBoxStyle.Information, "Information .....")
                        Exit Sub
                    End If
                End If
                i = i + 1
            Next
            OPR12.Visible = True
            UltraGroupBox36.Visible = False
            Call Load_GrideDye_Plan()
            Call Load_Gridewith_DyePlan()
            'UltraTabControl1.Tabs(4).Enabled = True
            'UltraTabControl1.SelectedTab = UltraTabControl1.Tabs(4)
                End If
    End Sub


    Function Load_Gridewith_DyePlan()
        Dim i As Integer

        Try


            i = 0
            For Each uRow As UltraGridRow In dg1.Rows
                If Trim(dg1.Rows(i).Cells(0).Text) <> "" Then
                    If Trim(dg1.Rows(i).Cells(9).Text) <> "" Then
                        Dim newRow As DataRow = c_dataCustomer5.NewRow


                        newRow("15Class") = Trim(dg1.Rows(i).Cells(0).Value)
                        newRow("Description") = Trim(dg1.Rows(i).Cells(1).Value)
                        ' newRow("Shade") = Trim(dg1.Rows(i).Cells(2).Value)
                        newRow("Allocate Yarn") = Trim(dg1.Rows(i).Cells(9).Value)
                        newRow("Allocate Con") = CInt(CDbl(Trim(dg1.Rows(i).Cells(9).Value)) / 1.05)
                        c_dataCustomer5.Rows.Add(newRow)
                    End If
                End If
                i = i + 1
            Next
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                'Con.close()
            End If
        End Try
    End Function

    Private Sub chkS_Yarn1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkS_Yarn1.CheckedChanged
        If chkS_Yarn1.Checked = True Then
            chkSp_Yarn2.Checked = False
        End If
    End Sub

    Private Sub chkSp_Yarn2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSp_Yarn2.CheckedChanged
        If chkSp_Yarn2.Checked = True Then
            chkS_Yarn1.Checked = False
        End If
    End Sub

    Function Load_Projection_Detailes(ByVal strDate As Date)
        Dim vcWhere As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim _Code As Integer
        Dim M01 As DataSet
        Dim M02 As DataSet

        Dim Value As Double
        Dim _VString As String
        Dim T01 As DataSet

        Try
            vcWhere = "select * from P01PARAMETER where P01CODE='PRN'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, vcWhere)
            If isValidDataset(M01) Then
                _Code = M01.Tables(0).Rows(0)("P01NO")
            End If

            _Code = _Code - 1
            i = 0
            vcWhere = "M43Quality='" & Trim(txtQuality.Text) & "'  and M43Count_No=" & _Code & " and M43Product_Month<=" & Month(strDate) & " and M43Year<=" & Year(strDate) & ""
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "PRS1"), New SqlParameter("@vcWhereClause1", vcWhere))
            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer2_KNT.NewRow

                newRow("##") = M01.Tables(0).Rows(i)("Code") & "|" & M01.Tables(0).Rows(i)("M43Product_Month") & "|" & M01.Tables(0).Rows(i)("M43Year")
                newRow("Shade") = M01.Tables(0).Rows(i)("M43Shade")
                'newRow("Month") = MonthName(M01.Tables(0).Rows(i)("M43Product_Month"))

                'newRow("Year") = M01.Tables(0).Rows(i)("M43Year")

                Value = M01.Tables(0).Rows(i)("Qty")
                _VString = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _VString = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Projection") = _VString

                'CHECK AVAILABLE ALLOCATED QTY FOR PROJECTION TMPPROJECTION TABLE
                'DEVELOPED BY SURANGA WIJESINGHE ON 01/08/2016
                Value = 0
                vcWhere = "tmpCode='" & Trim(M01.Tables(0).Rows(i)("Code")) & "' and tmpShade='" & Trim(M01.Tables(0).Rows(i)("M43Shade")) & "' and tmpQuality='" & txtQuality.Text & "' and tmpYear=" & M01.Tables(0).Rows(i)("M43Year") & " and tmpMonth=" & M01.Tables(0).Rows(i)("M43Product_Month") & ""
                T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TMP1"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(T01) Then
                    Value = T01.Tables(0).Rows(0)("tmpQty")
                End If

                'CHECK AVAILABLE ALLOCATED QTY FOR PROJECTION T15Projection
                'DEVELOPED BY SURANGA WIJESINGHE ON 01/08/2016

                vcWhere = "T15Code='" & Trim(M01.Tables(0).Rows(i)("Code")) & "' and T15Shade='" & Trim(M01.Tables(0).Rows(i)("M43Shade")) & "' and T15Quality='" & txtQuality.Text & "' and T15Year=" & M01.Tables(0).Rows(i)("M43Year") & " and T15Month=" & M01.Tables(0).Rows(i)("M43Product_Month") & ""
                T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TMP2"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(T01) Then
                    Value = Value + T01.Tables(0).Rows(0)("T15Qty")
                End If


                _VString = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _VString = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Allocated") = _VString

                'Allocation for this Line Item
                Value = 0
                vcWhere = "T15Code='" & Trim(M01.Tables(0).Rows(i)("Code")) & "' and T15Shade='" & Trim(M01.Tables(0).Rows(i)("M43Shade")) & "' and T15Quality='" & txtQuality.Text & "' and T15Year=" & M01.Tables(0).Rows(i)("M43Year") & " and T15Month=" & M01.Tables(0).Rows(i)("M43Product_Month") & " and T15Sales_Order='" & txtSales_Order.Text & "' and T15Line_Item=" & txtLine_Item.Text & ""
                T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TMP2"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(T01) Then
                    Value = Value + T01.Tables(0).Rows(0)("T15Qty")
                End If
                _VString = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _VString = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("For Quality") = _VString

                c_dataCustomer1_KNT.Rows.Add(newRow)


                i = i + 1
            Next

            con.close()
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Function Search_WeekNo()
        On Error Resume Next
        Dim _Date As Date
        Dim _TimeSpan As TimeSpan
        Dim userDate As Date
        Dim _LastDate As Date
        Dim _WeekNo As Integer

        userDate = "1/1/" & Year(txtComplete_Date_Knt.Text)
        If WeekdayName(Weekday(userDate)) = "Sunday" Then
            userDate = userDate.AddDays(-3)
        ElseIf WeekdayName(Weekday(userDate)) = "Monday" Then
            userDate = userDate.AddDays(-4)
        ElseIf WeekdayName(Weekday(userDate)) = "Tuesday" Then
            userDate = userDate.AddDays(-5)
        ElseIf WeekdayName(Weekday(userDate)) = "Thusday" Then
            'userDate = userDate.AddDays(-1)
        ElseIf WeekdayName(Weekday(userDate)) = "Friday" Then
            userDate = userDate.AddDays(-1)
        ElseIf WeekdayName(Weekday(userDate)) = "Saturday" Then
            userDate = userDate.AddDays(-2)

        End If
        _LastDate = txtComplete_Date_Knt.Text
        If WeekdayName(Weekday(_LastDate)) = "Sunday" Then
            _LastDate = _LastDate.AddDays(-3)
        ElseIf WeekdayName(Weekday(_LastDate)) = "Monday" Then
            _LastDate = _LastDate.AddDays(-4)
        ElseIf WeekdayName(Weekday(_LastDate)) = "Tuesday" Then
            _LastDate = _LastDate.AddDays(-5)
        ElseIf WeekdayName(Weekday(_LastDate)) = "Thusday" Then
            'userDate = userDate.AddDays(-1)
        ElseIf WeekdayName(Weekday(_LastDate)) = "Friday" Then
            _LastDate = _LastDate.AddDays(-1)
        ElseIf WeekdayName(Weekday(_LastDate)) = "Saturday" Then
            _LastDate = _LastDate.AddDays(-2)

        End If

        _TimeSpan = _LastDate.Subtract(userDate)
        _WeekNo = _TimeSpan.Days / 7

    End Function

    Private Sub dg1_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles dg1.ClickCellButton
        On Error Resume Next
        Dim _ColumIndex As Integer


        If UltraGroupBox16.Visible = True Then
            UltraGroupBox16.Visible = False
        Else
            _Rowindex = dg1.ActiveRow.Index
            _ColumIndex = dg1.ActiveCell.Column.Index

            '  If Trim(dg1.Rows(_Rowindex).Cells(0).Text) <> "" Then
            If _ColumIndex = 5 Then
                txt15Class.Text = dg1.Rows(_Rowindex).Cells(0).Value
                txtQty.Text = dg1.Rows(_Rowindex).Cells(5).Value
                If txt15Class.Text <> "" Then
                    UltraGroupBox16.Visible = True
                End If
                Call Load_Gride_StockCode()
                Call Load_Gride_DataStock()
            ElseIf _ColumIndex = 9 Then
                If UltraGroupBox20.Visible = True Then
                    UltraGroupBox20.Visible = False
                    With UltraGroupBox36
                        .Width = 546
                        .Height = 237
                    End With
                Else
                    txtY_Balance.Text = ""
                    txtY_Balance.Text = dg1.Rows(_Rowindex).Cells(7).Value
                    UltraGroupBox16.Visible = False
                    UltraGroupBox20.Visible = True

                    Call Load_Gride_YarnStock()
                    Call Load_GrideData_YarnStock()

                    With UltraGroupBox36
                        .Width = 546
                        .Height = 137
                    End With
                    ' 546, 237
                End If
            End If

            'End If
        End If
    End Sub

    Function Load_GrideData_YarnStock()
        Dim i As Integer
        Dim vcWhere As String
        Dim M01 As DataSet
        Dim M02 As DataSet
        Dim Value As Double
        Dim _VString As String
        Dim Diff As TimeSpan
        Dim _To As Date
        'Dim Value As Double
        Dim _Rowcount As Integer
        Dim _Date As Date

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim _String As String

        i = 0
        Try
            txtSearch.Text = ""
            _String = Trim(dg1.Rows(_Rowindex).Cells(1).Value)
            'If dg1.Rows(_Rowindex).Cells(0).Text <> "" Then
            vcWhere = "M33Yarn_Location='2020' and left(M33Description,4)='" & Microsoft.VisualBasic.Left(_String, 4) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "YST"), New SqlParameter("@vcWhereClause1", vcWhere))
            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                Value = 0
                Value = M01.Tables(0).Rows(i)("M33Qty")
                'T10Dyed_Yarn Table

                vcWhere = "T1015Class='" & txt15Class.Text & "' and T10Stock_Code='" & M01.Tables(0).Rows(i)("M33Stock_Code") & "'"
                M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "DYN"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M02) Then
                    Value = Value - M02.Tables(0).Rows(0)("Qty")
                End If

                vcWhere = "tmp10Class='" & M01.Tables(0).Rows(i)("M3310Class") & "' and tmpStock_Code='" & M01.Tables(0).Rows(i)("M33Stock_Code") & "' and tmp15Class<>'" & dg1.Rows(_Rowindex).Cells(0).Value & "'"
                M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "BTY1"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M02) Then
                    ' MsgBox("")
                    If IsDBNull(M02.Tables(0).Rows(0)("Qty")) Then
                    Else
                        Value = Value - M02.Tables(0).Rows(0)("Qty")
                    End If
                End If

                'tmpBlock_YarnStock
                vcWhere = "tmp10Class='" & M01.Tables(0).Rows(i)("M3310Class") & "' and tmpStock_Code='" & M01.Tables(0).Rows(i)("M33Stock_Code") & "' and tmpUser<>'" & strDisname & "'"
                M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "BTY"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M02) Then
                    Dim newRow As DataRow = c_dataCustomer4.NewRow

                    Dim _STValue As String

                    _STValue = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    _STValue = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                    newRow("##") = False
                    newRow("10 Class") = M01.Tables(0).Rows(i)("M3310Class")
                    newRow("Stock Code") = M01.Tables(0).Rows(i)("M33Stock_Code")
                    newRow("Description") = M01.Tables(0).Rows(i)("M33Description")
                    _Date = Month(M01.Tables(0).Rows(i)("M33Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("M33Date")) & "/" & Year(M01.Tables(0).Rows(i)("M33Date"))
                    Diff = Today.Subtract(_Date)
                    newRow("Age") = Diff.Days
                    newRow("Qty") = _STValue
                    newRow("Log User") = M02.Tables(0).Rows(0)("tmpUser")

                    c_dataCustomer4.Rows.Add(newRow)



                Else

                    Dim newRow As DataRow = c_dataCustomer4.NewRow

                    Dim _STValue As String

                    _STValue = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    _STValue = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                    newRow("##") = False
                    newRow("10 Class") = M01.Tables(0).Rows(i)("M3310Class")
                    newRow("Stock Code") = M01.Tables(0).Rows(i)("M33Stock_Code")
                    newRow("Description") = M01.Tables(0).Rows(i)("M33Description")
                    _Date = Month(M01.Tables(0).Rows(i)("M33Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("M33Date")) & "/" & Year(M01.Tables(0).Rows(i)("M33Date"))
                    Diff = Today.Subtract(_Date)
                    newRow("Age") = Diff.Days
                    newRow("Qty") = _STValue
                    newRow("Log User") = "-"

                    c_dataCustomer4.Rows.Add(newRow)


                End If

                i = i + 1
            Next
            Dim newRow1 As DataRow = c_dataCustomer4.NewRow
            c_dataCustomer4.Rows.Add(newRow1)

            ' End If
            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Function Load_Gride_DataStock()
        Dim i As Integer
        Dim vcWhere As String
        Dim M01 As DataSet
        Dim M02 As DataSet
        Dim Value As Double
        Dim _VString As String
        Dim Diff As TimeSpan
        Dim _To As Date
        'Dim Value As Double
        Dim _Rowcount As Integer

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        i = 0
        Try

            If dg1.Rows(_Rowindex).Cells(0).Text <> "" Then
                vcWhere = "M3310Class='" & Trim(dg1.Rows(_Rowindex).Cells(0).Value) & "' "
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "YST"), New SqlParameter("@vcWhereClause1", vcWhere))
                For Each DTRow3 As DataRow In M01.Tables(0).Rows
                    Value = 0
                    Value = M01.Tables(0).Rows(i)("M33Qty")
                    'T10Dyed_Yarn Table

                    vcWhere = "T1015Class='" & txt15Class.Text & "' and T10Stock_Code='" & M01.Tables(0).Rows(i)("M33Stock_Code") & "'"
                    M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "DYN"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M02) Then
                        Value = Value - M02.Tables(0).Rows(0)("Qty")
                    End If

                    'tmpBlock_YarnStock
                    vcWhere = "tmp15Class='" & txt15Class.Text & "' and tmpStock_Code='" & M01.Tables(0).Rows(i)("M33Stock_Code") & "' and tmpUser<>'" & strDisname & "'"
                    M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "BTY"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M02) Then
                        Dim newRow As DataRow = c_dataCustomer3.NewRow

                        Dim _STValue As String

                        _STValue = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        _STValue = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                        newRow("##") = False
                        newRow("Stock Code") = M01.Tables(0).Rows(i)("M33Stock_Code")
                        newRow("Description") = M01.Tables(0).Rows(i)("M33Description")
                        newRow("Date") = Month(M01.Tables(0).Rows(i)("M33Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("M33Date")) & "/" & Year(M01.Tables(0).Rows(i)("M33Date"))

                        newRow("Qty") = _STValue
                        newRow("Log User") = M02.Tables(0).Rows(0)("tmpUser")

                        c_dataCustomer3.Rows.Add(newRow)



                    Else

                        Dim newRow As DataRow = c_dataCustomer3.NewRow

                        Dim _STValue As String

                        _STValue = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        _STValue = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                        newRow("##") = False
                        newRow("Stock Code") = M01.Tables(0).Rows(i)("M33Stock_Code")
                        newRow("Description") = M01.Tables(0).Rows(i)("M33Description")
                        newRow("Date") = Month(M01.Tables(0).Rows(i)("M33Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("M33Date")) & "/" & Year(M01.Tables(0).Rows(i)("M33Date"))
                        newRow("Qty") = _STValue
                        newRow("Log User") = "-"

                        c_dataCustomer3.Rows.Add(newRow)


                    End If

                    i = i + 1
                Next
                Dim newRow1 As DataRow = c_dataCustomer1.NewRow
                c_dataCustomer1.Rows.Add(newRow1)

            End If
            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                Con.close()
            End If
        End Try
    End Function

    Private Sub dg2_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles dg2.AfterRowUpdate

        Dim _Index As Integer
        Dim i As Integer
        Dim Value As Double
        Dim _stBalance As String

        Try
            _Index = dg2.ActiveRow.Index
            If dg2.Rows(_Index).Cells(1).Text <> "" Then
                ' MsgBox(Trim(UltraGrid2.Rows(_Index).Cells(0).Text))
                If dg2.Rows(_Index).Cells(0).Text = True Then

                    If CDbl(dg2.Rows(_Index).Cells(4).Value) <= txtQty.Text And Trim(UltraGrid2.Rows(_Index).Cells(6).Text) <> strDisname Then
                        dg2.Rows(_Index).Cells(5).Value = dg2.Rows(_Index).Cells(4).Value
                    End If
                Else
                    ' UltraGrid2.Rows(_Index).Cells(5).Value = ""
                End If
                If Trim(dg2.Rows(_Index).Cells(6).Value) <> "" Then

                    If Trim(dg2.Rows(_Index).Cells(6).Value) = "-" Then
                        '  Call Update_Transaction(Trim(txt15Class.Text), UltraGrid2.Rows(_Index).Cells(1).Value, UltraGrid2.Rows(_Index).Cells(5).Value)

                        Value = 0
                        _stBalance = ""
                        i = 0
                        For Each uRow As UltraGridRow In dg2.Rows
                            If IsNumeric(dg2.Rows(i).Cells(5).Value) Then
                                Value = Value + dg2.Rows(i).Cells(5).Value
                            End If
                            i = i + 1
                        Next

                        If Value <= txtQty.Text Then
                            Call Update_Transaction(Trim(txt15Class.Text), dg2.Rows(_Index).Cells(1).Value, dg2.Rows(_Index).Cells(5).Value)

                            _stBalance = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                            _stBalance = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                            dg1.Rows(_Rowindex).Cells(4).Value = _stBalance

                            Value = CDbl(txtQty.Text) - Value
                            Value = Value + (Value * (txtYarn_Wst.Text / 100))
                            _stBalance = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                            _stBalance = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                            dg1.Rows(_Rowindex).Cells(6).Value = _stBalance

                            dg1.Rows(_Rowindex).Cells(7).Value = CInt(Value / 1.05)
                        Else
                            MsgBox("Stock Quantity miss match please try again", MsgBoxStyle.Exclamation, "Technova ......")
                            Exit Sub
                        End If
                    End If
                End If
            End If

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)

            End If
        End Try
    End Sub

    Function Delete_Transaction()
        Dim vcWhere As String
        Dim M01 As DataSet
        Dim vcFieldList As String
        Dim ncQryType As String
        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Try

            nvcFieldList1 = "select * from T01Delivary_Request  where T01Sales_Order='" & strSales_Order & "' and T01Line_Item='" & strLine_Item & "'"
            M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(M01) Then
                txtreq_Date.Text = M01.Tables(0).Rows(0)("T01RQD")
                txtT2_req.Text = M01.Tables(0).Rows(0)("T01RQD")
                txtT4_req.Text = M01.Tables(0).Rows(0)("T01RQD")
            End If

            nvcFieldList1 = "DELETE FROM tmpBlock_Yarn_Stock_Code WHERE tmpRefNo=" & Delivary_Ref & " "
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            nvcFieldList1 = "DELETE FROM tmpYarn_Booking WHERE tmpRef=" & Delivary_Ref & " "
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            transaction.Commit()
            connection.Close()
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)

                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
                connection.Close()
            End If
        End Try


    End Function

    Function Update_Transaction(ByVal str15 As String, ByVal strStock As String, ByVal strQty As Double)
        Dim vcWhere As String
        Dim M01 As DataSet
        Dim vcFieldList As String
        Dim ncQryType As String
        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Try
            ''nvcFieldList1 = "DELETE FROM tmpBlock_Yarn_Stock_Code WHERE tmpRefNo=" & Delivary_Ref & " "
            ''ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            vcWhere = "tmp15Class='" & dg1.Rows(_Rowindex).Cells(0).Value & "' and tmpStock_Code='" & strStock & "' and tmpUser='" & strDisname & "' and tmp10Class='" & str15 & "'"
            M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "YSS"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                If strDisname = Trim(M01.Tables(0).Rows(0)("tmpUser")) Then
                    nvcFieldList1 = "update tmpBlock_Yarn_Stock_Code set tmpQty='" & strQty & "' where tmp15Class='" & dg1.Rows(_Rowindex).Cells(0).Value & "' and tmpStock_Code='" & strStock & "' and tmpUser='" & strDisname & "' and tmp10Class='" & str15 & "' AND tmpRefNo=" & Delivary_Ref & ""
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                End If
            Else

                ncQryType = "ADD1"
                nvcFieldList1 = "(tmpRefNo," & "tmp15Class," & "tmpStock_Code," & "tmpQty," & "tmpUser," & "tmp10Class," & "tmpDis) " & "values(" & Delivary_Ref & ",'" & dg1.Rows(_Rowindex).Cells(0).Value & "','" & strStock & "','" & strQty & "','" & strDisname & "','" & str15 & "','" & dg1.Rows(_Rowindex).Cells(1).Value & "')"
                up_GetSetDelivary_Planning(ncQryType, nvcFieldList1, vcWhere, connection, transaction)
            End If

            'INSERT TEMYARN_BOOKING TABLE ON 2016.4.25
            vcWhere = "tmp10Class='" & dg1.Rows(_Rowindex).Cells(0).Value & "' and tmpStock_Code='" & strStock & "' and tmpSO='" & strSales_Order & "' and tmpLine_Item=" & strLine_Item & ""
            M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "YARN"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                nvcFieldList1 = "update tmpYarn_Booking set tmpQty='" & strQty & "' where  tmp10Class='" & dg1.Rows(_Rowindex).Cells(0).Value & "' and tmpStock_Code='" & strStock & "' and tmpSO='" & strSales_Order & "' and tmpLine_Item=" & strLine_Item & ""
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            Else
                ncQryType = "YADD"
                nvcFieldList1 = "(tmpRef," & "tmp10Class," & "tmpStock_Code," & "tmpLocation," & "tmpQty," & "tmpStatus," & "tmpCat," & "tmpSO," & "tmpLine_Item," & "tmpTime," & "tmpB_Status) " & "values(" & Delivary_Ref & ",'" & dg1.Rows(_Rowindex).Cells(0).Value & "','" & strStock & "','2020','" & strQty & "','A','-','" & strSales_Order & "'," & strLine_Item & ",'" & Now & "','YD')"
                up_GetSetYarn_Bookingtmp(ncQryType, nvcFieldList1, vcWhere, connection, transaction)
            End If
            transaction.Commit()
            connection.Close()
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)

                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
                connection.Close()
            End If
        End Try


    End Function

    Function Calculate_Yarn_Qty(ByVal _GdIdex As Integer)
        Dim _Index As Integer
        Dim i As Integer
        Dim Value As Double
        Dim _stBalance As String
        Dim _String As String
        Dim _Qty As Double
        Dim _BalanceQty As Double
        Try
            _BalanceQty = CDbl(txtY_Balance.Text)
            i = 0
            For Each uRow As UltraGridRow In dg1.Rows
                If IsNumeric(dg1.Rows(i).Cells(9).Value) Then
                    _Qty = _Qty + dg1.Rows(i).Cells(9).Value
                End If
                i = i + 1
            Next

            If _BalanceQty >= _Qty Then
                _BalanceQty = _BalanceQty - _Qty
                If _BalanceQty > 0 Then
                    If dg4.Rows(_GdIdex).Cells(1).Text <> "" Then
                        If dg4.Rows(_GdIdex).Cells(0).Text = True Then
                            If dg4.Rows(_GdIdex).Cells(4).Value >= _BalanceQty Then
                                dg4.Rows(_GdIdex).Cells(6).Value = _BalanceQty
                            Else
                                dg4.Rows(_GdIdex).Cells(6).Value = dg4.Rows(_GdIdex).Cells(4).Value
                            End If
                        Else
                            dg4.Rows(_GdIdex).Cells(6).Value = ""
                        End If
                        'CALCULATING QTY
                        i = 0
                        _Qty = 0

                        'DELETE TMPYARN_BOOKING TABLE AND TMP_BLOCK_YARN_STOCK_CODE TABLE
                        Call Delete_Transaction_Yarn_Booking()

                        For Each uRow As UltraGridRow In dg1.Rows
                            Dim _10Class As String
                            _10Class = Trim(dg4.Rows(i).Cells(1).Value)
                            If IsNumeric(dg4.Rows(i).Cells(6).Value) Then
                                _Qty = _Qty + dg4.Rows(i).Cells(6).Value
                                'UPDATE RECORDS
                                Call Update_Transaction(Trim(_10Class), dg4.Rows(i).Cells(3).Value, dg4.Rows(i).Cells(6).Value)
                            End If
                            i = i + 1
                        Next

                        If _Qty > 0 Then
                            _stBalance = (_Qty.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                            _stBalance = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", _Qty))

                            dg1.Rows(_Rowindex).Cells(9).Value = _stBalance



                        End If
                    End If
                End If

            Else
                Exit Function
            End If

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)

            End If
        End Try
    End Function
    Private Sub dg4_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles dg4.AfterRowUpdate
        Dim _Index As Integer
        Dim i As Integer
        Dim Value As Double
        Dim _stBalance As String
        Dim _String As String
        Dim _Qty As Double
        '  Try
        _Index = dg4.ActiveRow.Index

        Call Calculate_Yarn_Qty(_Index)
        ' ''    If dg4.Rows(_Index).Cells(1).Text <> "" Then
        ' ''        ' MsgBox(Trim(UltraGrid2.Rows(_Index).Cells(0).Text))
        ' ''        If dg4.Rows(_Index).Cells(0).Text = True Then
        ' ''            i = 0

        ' ''            If dg4.Rows(_Index).Cells(4).Value <= txtBalance.Text And Trim(dg4.Rows(_Index).Cells(7).Text) <> strDisname Then
        ' ''                dg4.Rows(_Index).Cells(6).Value = dg4.Rows(_Index).Cells(4).Value
        ' ''            End If
        ' ''        Else
        ' ''            dg4.Rows(_Index).Cells(6).Value = ""
        ' ''        End If

        ' ''        If Trim(dg4.Rows(_Index).Cells(6).Text) <> "" Then
        ' ''            Dim _10Class As String
        ' ''            _10Class = Trim(dg4.Rows(_Index).Cells(1).Value)
        ' ''            If Trim(dg4.Rows(_Index).Cells(7).Value) = "-" Then
        ' ''                '  Call Update_Transaction(Trim(_10Class), UltraGrid3.Rows(_Index).Cells(3).Value, UltraGrid3.Rows(_Index).Cells(6).Value)

        ' ''                Value = 0
        ' ''                _stBalance = ""
        ' ''                i = 0
        ' ''                For Each uRow As UltraGridRow In dg4.Rows
        ' ''                    If IsNumeric(dg4.Rows(i).Cells(6).Value) Then
        ' ''                        Value = Value + dg4.Rows(i).Cells(6).Value
        ' ''                    End If
        ' ''                    i = i + 1
        ' ''                Next

        ' ''                If Value <= txtBalance.Text Then
        ' ''                    Call Update_Transaction(Trim(_10Class), dg4.Rows(_Index).Cells(3).Value, dg4.Rows(_Index).Cells(6).Value)
        ' ''                    _stBalance = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
        ' ''                    _stBalance = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

        ' ''                    dg1.Rows(_Rowindex).Cells(9).Value = _stBalance

        ' ''                    Value = _stBalance
        ' ''                    '  Value = Value + (Value * (txtYarn_Wst.Text / 100))
        ' ''                    _stBalance = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
        ' ''                    _stBalance = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

        ' ''                    dg1.Rows(_Rowindex).Cells(9).Value = _stBalance


        ' ''                    'If txtDis.Text <> "" Then

        ' ''                    '    'txtDis.Focus()

        ' ''                    '    'SendKeys.Send("{ENTER}")
        ' ''                    '    _String = txtDis.Text & ";" & UltraGrid1.Rows(_Rowindex).Cells(0).Value & "- Allocate Cons(" & CInt(CDbl(_stBalance) / 1.05) & ")"
        ' ''                    '    txtDis.Text = _String
        ' ''                    '    'Dim Words As String() = _String.Split(New Char() {";"c})
        ' ''                    '    'txtDis.Text = UltraGrid1.Rows(_Rowindex).Cells(0).Value & "- Allocate Cons(" & CInt(CDbl(_stBalance) / 1.05) & ")"

        ' ''                    '    '                                txtDis.Text = Words(0)

        ' ''                    '    'SendKeys.Send("{ENTER}")
        ' ''                    '    '                                txtDis.Text = Words(1)
        ' ''                    'Else
        ' ''                    '    txtDis.Text = UltraGrid1.Rows(_Rowindex).Cells(0).Value & "- Allocate Cons(" & CInt(CDbl(_stBalance) / 1.05) & ")"
        ' ''                    '    _String = txtDis.Text
        ' ''                    'End If
        ' ''                    '  UltraGrid1.Rows(_Rowindex).Cells(7).Value = CInt(Value / 1.05)
        ' ''                Else
        ' ''                    MsgBox("Stock Quantity miss match please try again", MsgBoxStyle.Exclamation, "Technova ......")
        ' ''                    Exit Sub
        ' ''                End If
        ' ''            End If
        ' ''        End If
        ' ''    End If

        ' ''Catch returnMessage As EvaluateException
        ' ''    If returnMessage.Message <> Nothing Then
        ' ''        MessageBox.Show(returnMessage.Message)

        ' ''    End If
        ' ''End Try
    End Sub


    Private Sub cmdSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSearch.Click
        Call Load_Gride_YarnStock()
        Call Load_GrideData_YarnStockLike()
    End Sub

    Function Load_GrideData_YarnStockLike()
        Dim i As Integer
        Dim vcWhere As String
        Dim M01 As DataSet
        Dim M02 As DataSet
        Dim Value As Double
        Dim _VString As String
        Dim Diff As TimeSpan
        Dim _To As Date
        'Dim Value As Double
        Dim _Rowcount As Integer
        Dim _Date As Date

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        i = 0
        Try

            If dg1.Rows(_Rowindex).Cells(0).Text <> "" Then
                vcWhere = "M33Yarn_Location='2020' and M33Description like '%" & txtSearch.Text & "%' and left(M33Description,4)='" & Microsoft.VisualBasic.Left(txtBasic_Yarn.Text, 4) & "'"
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "YST"), New SqlParameter("@vcWhereClause1", vcWhere))
                For Each DTRow3 As DataRow In M01.Tables(0).Rows
                    Value = 0
                    Value = M01.Tables(0).Rows(i)("M33Qty")
                    'T10Dyed_Yarn Table

                    vcWhere = "T1015Class='" & txt15Class.Text & "' and T10Stock_Code='" & M01.Tables(0).Rows(i)("M33Stock_Code") & "'"
                    M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "DYN"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M02) Then
                        Value = Value - M02.Tables(0).Rows(0)("Qty")
                    End If

                    'tmpBlock_YarnStock
                    vcWhere = "tmp15Class='" & M01.Tables(0).Rows(i)("M3310Class") & "' and tmpStock_Code='" & M01.Tables(0).Rows(i)("M33Stock_Code") & "' and tmpUser<>'" & strDisname & "'"
                    M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "BTY"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M02) Then
                        Dim newRow As DataRow = c_dataCustomer5.NewRow

                        Dim _STValue As String

                        _STValue = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        _STValue = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                        newRow("##") = False
                        newRow("10 Class") = M01.Tables(0).Rows(i)("M3310Class")
                        newRow("Stock Code") = M01.Tables(0).Rows(i)("M33Stock_Code")
                        newRow("Description") = M01.Tables(0).Rows(i)("M33Description")
                        _Date = Month(M01.Tables(0).Rows(i)("M33Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("M33Date")) & "/" & Year(M01.Tables(0).Rows(i)("M33Date"))
                        Diff = Today.Subtract(_Date)
                        newRow("Age") = Diff.Days
                        newRow("Qty") = _STValue
                        newRow("Log User") = M02.Tables(0).Rows(0)("tmpUser")

                        c_dataCustomer5.Rows.Add(newRow)



                    Else

                        Dim newRow As DataRow = c_dataCustomer5.NewRow

                        Dim _STValue As String

                        _STValue = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        _STValue = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                        newRow("##") = False
                        newRow("10 Class") = M01.Tables(0).Rows(i)("M3310Class")
                        newRow("Stock Code") = M01.Tables(0).Rows(i)("M33Stock_Code")
                        newRow("Description") = M01.Tables(0).Rows(i)("M33Description")
                        _Date = Month(M01.Tables(0).Rows(i)("M33Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("M33Date")) & "/" & Year(M01.Tables(0).Rows(i)("M33Date"))
                        Diff = Today.Subtract(_Date)
                        newRow("Age") = Diff.Days
                        newRow("Qty") = _STValue
                        newRow("Log User") = "-"

                        c_dataCustomer5.Rows.Add(newRow)


                    End If

                    i = i + 1
                Next
                Dim newRow1 As DataRow = c_dataCustomer5.NewRow
                c_dataCustomer5.Rows.Add(newRow1)

            End If
            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Private Sub UltraButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton4.Click
        Dim vcWhere As String
        Dim M01 As DataSet
        Dim vcFieldList As String
        Dim ncQryType As String
        Dim nvcFieldList1 As String
        Dim M02 As DataSet

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Dim _MC As String
        Dim _string As String
        Dim result1 As DialogResult
        Dim _Date As Date
        Dim i As Integer
        Dim _DyeStartDate As Date
        Dim _DyeEnd_Date As Date
        Dim _Balance_Qty As Double
        Dim X As Integer
        Dim _Status As Boolean
        Dim _No_Of_Batch As Integer

        Try
            _Date = txtDate.Text
            If txtWeek.Text <> "" Then
            Else
                txtYear.Text = Year(txtDate.Text)
                txtWeek.Text = DatePart("WW", _Date, FirstDayOfWeek.Thursday)
            End If

            If IsNumeric(txtWeek.Text) Then
            Else
                MsgBox("Please enter the correct Week No", MsgBoxStyle.Information, "Information ......")
                txtWeek.Focus()
                Exit Sub
            End If

            If IsNumeric(txtYear.Text) Then
            Else
                MsgBox("Please enter the correct Year", MsgBoxStyle.Information, "Information ......")
                txtYear.Focus()
                Exit Sub
            End If
            Dim _WeekDel_Date As Date

            If Trim(txtWeek.Text) <> "" Then
                If Trim(txtYear.Text) <> "" Then
                Else
                    MsgBox("Please enter the Year", MsgBoxStyle.Information, "Information .......")
                    Exit Sub
                End If
                Dim StartDate As New DateTime(txtYear.Text, 1, 1)
                _WeekDel_Date = DateAdd(DateInterval.WeekOfYear, txtWeek.Text - 1, StartDate)
                ' MsgBox(WeekdayName(Weekday(_WeekDel_Date)))
                If (WeekdayName(Weekday(_WeekDel_Date))) = "Sunday" Then
                    _WeekDel_Date = _WeekDel_Date.AddDays(+4)
                ElseIf (WeekdayName(Weekday(_WeekDel_Date))) = "Monday" Then
                    _WeekDel_Date = _WeekDel_Date.AddDays(+3)
                ElseIf (WeekdayName(Weekday(_WeekDel_Date))) = "Tuesday" Then
                    _WeekDel_Date = _WeekDel_Date.AddDays(+2)


                ElseIf (WeekdayName(Weekday(_WeekDel_Date))) = "Wednesday" Then
                    _WeekDel_Date = _WeekDel_Date.AddDays(+1)
                ElseIf (WeekdayName(Weekday(_WeekDel_Date))) = "Friday" Then
                    _WeekDel_Date = _WeekDel_Date.AddDays(-1)
                ElseIf (WeekdayName(Weekday(_WeekDel_Date))) = "Saturday" Then
                    _WeekDel_Date = _WeekDel_Date.AddDays(-1)
                End If

            Else
                If txtDate.Text > Today Then
                    _WeekDel_Date = txtDate.Text
                Else
                    MsgBox("Please check the delivary date", MsgBoxStyle.Information, "Information .....")
                    txtDate.Focus()
                    Exit Sub
                End If


            End If

            'Check Dye MC Block
            M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "DMC"))
            If isValidDataset(M01) Then
                If Trim(M01.Tables(0).Rows(0)("M37User")) = strDisname Then
                Else
                    result1 = MessageBox.Show(M01.Tables(0).Rows(0)("M37User") & " used this dye machine.", _
                                    "Error ....", _
                                    MessageBoxButtons.OK, MessageBoxIcon.Error)
                    If result1 = Windows.Forms.DialogResult.OK Then
                        Exit Sub
                    End If
                End If
            Else
                ncQryType = "ADD"
                nvcFieldList1 = "(M37Date," & "M37User) " & "values('" & Today & "','" & strDisname & "')"
                up_GetSetBlock_DyeMC(ncQryType, nvcFieldList1, vcWhere, connection, transaction)
            End If

            transaction.Commit()
            connection.Close()
            '-------------------------------------------------------------------------------------------------
            ' Call Load_GrideDye_Plan()
            ' Call Load_GrideDye_Plan()
            i = 0
            _DyeStartDate = _Date.AddDays(-14)

            'For Each uRow As UltraGridRow In dg1.Rows
            '    If UltraGrid1.Rows(i).Cells(0).Text <> "" Then
            '        Dim newRow1 As DataRow = c_dataCustomer2.NewRow



            '        newRow1("15Class") = dg1.Rows(i).Cells(0).Value
            '        newRow1("Description") = dg1.Rows(i).Cells(1).Value

            '        c_dataCustomer2.Rows.Add(newRow1)
            '    Else
            '        i = i + 1
            '        Continue For
            '    End If
            '    i = i + 1
            'Next

            'Dim newRow As DataRow = c_dataCustomer2.NewRow
            'newRow("15Class") = ""
            'newRow("Description") = ""
            'c_dataCustomer2.Rows.Add(newRow)
            '----------------------------------------------------------------------------
            'Check the Suterble Machine
            Dim _row As Integer
            _Balance_Qty = 0
            i = 0
            _Status = False
            _row = 0
            connection = DBEngin.GetConnection(True)
            connectionCreated = True
            transaction = connection.BeginTransaction()
            transactionCreated = True

            For Each uRow As UltraGridRow In dg1.Rows
                If Trim(dg1.Rows(i).Cells(0).Text) <> "" Then
                    _Balance_Qty = CInt(dg1.Rows(i).Cells(7).Value) + 5
                    vcFieldList = "left(m36mc_no,1)='Y' and m36Max_Qty<='" & _Balance_Qty & "'"
                    M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDyeMC_Process", New SqlParameter("@cQryType", "MCC"), New SqlParameter("@vcWhereClause1", vcFieldList))
                    X = 0
                    _No_Of_Batch = 0
                    For Each DTRow3 As DataRow In M01.Tables(0).Rows
                        _DyeStartDate = _Date.AddDays(-14)
                        If M01.Tables(0).Rows(X)("m36min_qty") <= _Balance_Qty And M01.Tables(0).Rows(X)("m36max_qty") >= _Balance_Qty Then
                            vcFieldList = "tmpMC_No='" & Trim(M01.Tables(0).Rows(X)("M36MC_No")) & "'"
                            M02 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDyeMC_Process", New SqlParameter("@cQryType", "DYM"), New SqlParameter("@vcWhereClause1", vcFieldList))
                            If isValidDataset(M02) Then
                                ' MsgBox(CDate(M02.Tables(0).Rows(X)("tmpEnd_Time")))
                                If CDate(M02.Tables(0).Rows(0)("tmpEnd_Time")) <= _DyeStartDate Then
                                    _No_Of_Batch = _No_Of_Batch + (M01.Tables(0).Rows(X)("m36max_qty") / _Balance_Qty)
                                    ' MsgBox(CDate(M02.Tables(0).Rows(X)("tmpEnd_Time")))
                                    _DyeEnd_Date = CDate(M02.Tables(0).Rows(0)("tmpEnd_Time")).AddHours(+10 * _No_Of_Batch)
                                    _DyeStartDate = M02.Tables(0).Rows(0)("tmpEnd_Time")
                                    If CDate(_DyeEnd_Date) > txtDate.Text Then

                                    Else
                                        _Status = True
                                        dg3.Rows(i).Cells(2).Value = _No_Of_Batch
                                        dg3.Rows(i).Cells(3).Value = _Balance_Qty
                                        dg3.Rows(i).Cells(4).Value = CInt(_Balance_Qty / 1.05)
                                        dg3.Rows(i).Cells(5).Value = M01.Tables(0).Rows(X)("M36MC_No")
                                        dg1.Rows(i).Cells(7).Value = UltraGrid1.Rows(i).Cells(7).Value + (5 * _No_Of_Batch)

                                        ncQryType = "ADD1"
                                        nvcFieldList1 = "(tmpRefNo," & "tmpMC_No," & "tmp15Class," & "tmpDate," & "tmpSTTime," & "tmpEnd_Time," & "tmpQty," & "tmpStatus," & "tmpShade," & "tmpCatagary," & "tmpCon) " & "values('" & Delivary_Ref & "','" & Trim(M01.Tables(0).Rows(X)("M36MC_No")) & "','" & dg1.Rows(i).Cells(0).Value & "','" & CDate(_DyeEnd_Date) & "','" & _DyeStartDate & "','" & _DyeEnd_Date & "','" & _Balance_Qty & "','I','" & UltraGrid1.Rows(i).Cells(2).Value & "','" & dg1.Rows(i).Cells(3).Value & "','" & CInt(_Balance_Qty / 1.05) & "')"
                                        up_GetSetYarn_DyeMCPln(ncQryType, nvcFieldList1, vcWhere, connection, transaction)
                                        'transaction.Commit()
                                        Exit For
                                    End If
                                End If
                            Else
                                _Status = True
                                _DyeStartDate = _DyeStartDate & " 7:30AM"
                                _No_Of_Batch = _No_Of_Batch + (M01.Tables(0).Rows(X)("m36max_qty") / _Balance_Qty)
                                _DyeEnd_Date = _DyeStartDate.AddHours(+10)

                                dg3.Rows(i).Cells(2).Value = _No_Of_Batch
                                dg3.Rows(i).Cells(3).Value = _Balance_Qty
                                dg3.Rows(i).Cells(4).Value = CInt(_Balance_Qty / 1.05)
                                dg3.Rows(i).Cells(5).Value = M01.Tables(0).Rows(X)("M36MC_No")
                                dg1.Rows(i).Cells(7).Value = dg1.Rows(i).Cells(7).Value + (5 * _No_Of_Batch)

                                ncQryType = "ADD1"
                                nvcFieldList1 = "(tmpRefNo," & "tmpMC_No," & "tmp15Class," & "tmpDate," & "tmpSTTime," & "tmpEnd_Time," & "tmpQty," & "tmpStatus," & "tmpShade," & "tmpCatagary," & "tmpCon) " & "values('" & Delivary_Ref & "','" & Trim(M01.Tables(0).Rows(X)("M36MC_No")) & "','" & dg1.Rows(i).Cells(0).Value & "','" & CDate(_DyeEnd_Date) & "','" & _DyeStartDate & "','" & _DyeEnd_Date & "','" & _Balance_Qty & "','I','" & dg1.Rows(i).Cells(1).Value & "','" & dg1.Rows(i).Cells(2).Value & "','" & CInt(_Balance_Qty / 1.05) & "')"
                                up_GetSetYarn_DyeMCPln(ncQryType, nvcFieldList1, vcWhere, connection, transaction)
                                ' transaction.Commit()
                                Exit For
                            End If
                        ElseIf M01.Tables(0).Rows(X)("m36min_qty") >= _Balance_Qty Then
                            Continue For

                        ElseIf M01.Tables(0).Rows(X)("M36max_Qty") <= _Balance_Qty Then
                            vcFieldList = "tmpMC_No='" & Trim(M01.Tables(0).Rows(X)("M36MC_No")) & "'"
                            M02 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDyeMC_Process", New SqlParameter("@cQryType", "DYM"), New SqlParameter("@vcWhereClause1", vcFieldList))
                            If isValidDataset(M02) Then
                                ' MsgBox(CDate(M02.Tables(0).Rows(X)("tmpEnd_Time")))
                                If CDate(M02.Tables(0).Rows(0)("tmpEnd_Time")) <= _DyeStartDate Then
                                    _No_Of_Batch = _No_Of_Batch + (M01.Tables(0).Rows(X)("m36max_qty") / _Balance_Qty)
                                    ' MsgBox(CDate(M02.Tables(0).Rows(X)("tmpEnd_Time")))
                                    _DyeEnd_Date = CDate(M02.Tables(0).Rows(0)("tmpEnd_Time")).AddHours(+10 * _No_Of_Batch)
                                    _DyeStartDate = M02.Tables(0).Rows(0)("tmpEnd_Time")
                                    If CDate(_DyeEnd_Date) > txtDate.Text Then

                                    Else
                                        _Status = True
                                        dg3.Rows(i).Cells(2).Value = _No_Of_Batch
                                        dg3.Rows(i).Cells(3).Value = _Balance_Qty
                                        dg3.Rows(i).Cells(4).Value = CInt(_Balance_Qty / 1.05)
                                        dg3.Rows(i).Cells(5).Value = M01.Tables(0).Rows(X)("M36MC_No")
                                        dg1.Rows(i).Cells(7).Value = UltraGrid1.Rows(i).Cells(7).Value + (5 * _No_Of_Batch)

                                        ncQryType = "ADD1"
                                        nvcFieldList1 = "(tmpRefNo," & "tmpMC_No," & "tmp15Class," & "tmpDate," & "tmpSTTime," & "tmpEnd_Time," & "tmpQty," & "tmpStatus," & "tmpShade," & "tmpCatagary," & "tmpCon) " & "values('" & Delivary_Ref & "','" & Trim(M01.Tables(0).Rows(X)("M36MC_No")) & "','" & dg1.Rows(i).Cells(0).Value & "','" & CDate(_DyeEnd_Date) & "','" & _DyeStartDate & "','" & _DyeEnd_Date & "','" & _Balance_Qty & "','I','" & UltraGrid1.Rows(i).Cells(2).Value & "','" & dg1.Rows(i).Cells(3).Value & "','" & CInt(_Balance_Qty / 1.05) & "')"
                                        up_GetSetYarn_DyeMCPln(ncQryType, nvcFieldList1, vcWhere, connection, transaction)
                                        'transaction.Commit()
                                        Exit For
                                    End If
                                End If
                            Else
                                _Status = True
                                _DyeStartDate = _DyeStartDate & " 7:30AM"
                                _No_Of_Batch = _No_Of_Batch + (M01.Tables(0).Rows(X)("m36max_qty") / _Balance_Qty)
                                _DyeEnd_Date = _DyeStartDate.AddHours(+10)

                                dg3.Rows(i).Cells(2).Value = _No_Of_Batch
                                dg3.Rows(i).Cells(3).Value = _Balance_Qty
                                dg3.Rows(i).Cells(4).Value = CInt(_Balance_Qty / 1.05)
                                dg3.Rows(i).Cells(5).Value = M01.Tables(0).Rows(X)("M36MC_No")
                                dg1.Rows(i).Cells(7).Value = dg1.Rows(i).Cells(7).Value + (5 * _No_Of_Batch)

                                ncQryType = "ADD1"
                                nvcFieldList1 = "(tmpRefNo," & "tmpMC_No," & "tmp15Class," & "tmpDate," & "tmpSTTime," & "tmpEnd_Time," & "tmpQty," & "tmpStatus," & "tmpShade," & "tmpCatagary," & "tmpCon) " & "values('" & Delivary_Ref & "','" & Trim(M01.Tables(0).Rows(X)("M36MC_No")) & "','" & dg1.Rows(i).Cells(0).Value & "','" & CDate(_DyeEnd_Date) & "','" & _DyeStartDate & "','" & _DyeEnd_Date & "','" & _Balance_Qty & "','I','" & dg1.Rows(i).Cells(1).Value & "','" & dg1.Rows(i).Cells(2).Value & "','" & CInt(_Balance_Qty / 1.05) & "')"
                                up_GetSetYarn_DyeMCPln(ncQryType, nvcFieldList1, vcWhere, connection, transaction)
                                ' transaction.Commit()
                                Exit For
                            End If
                        End If
                        X = X + 1
                    Next
                Else
                    Dim _Machine_No As String
                    _Machine_No = ""
                    If _Status = False Then
                        vcFieldList = "left(m36mc_no,1)='Y' and m36Max_Qty<='" & _Balance_Qty & "'"
                        M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDyeMC_Process", New SqlParameter("@cQryType", "MCC"), New SqlParameter("@vcWhereClause1", vcFieldList))
                        X = 0
                        _No_Of_Batch = 0
                        For Each DTRow3 As DataRow In M01.Tables(0).Rows
                            _DyeStartDate = _Date.AddDays(-14)
                            If M01.Tables(0).Rows(X)("m36min_qty") <= _Balance_Qty Then
                                vcFieldList = "tmpMC_No='" & Trim(M01.Tables(0).Rows(X)("M36MC_No")) & "'"
                                M02 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDyeMC_Process", New SqlParameter("@cQryType", "DYM"), New SqlParameter("@vcWhereClause1", vcFieldList))
                                If isValidDataset(M02) Then
                                    ' MsgBox(CDate(M02.Tables(0).Rows(X)("tmpEnd_Time")))
                                    If CDate(M02.Tables(0).Rows(0)("tmpEnd_Time")) <= _DyeStartDate Then
                                        _No_Of_Batch = _No_Of_Batch + (M01.Tables(0).Rows(X)("m36max_qty") / _Balance_Qty)
                                        ' MsgBox(CDate(M02.Tables(0).Rows(X)("tmpEnd_Time")))
                                        _DyeEnd_Date = CDate(M02.Tables(0).Rows(0)("tmpEnd_Time")).AddHours(+10 * _No_Of_Batch)
                                        _DyeStartDate = M02.Tables(0).Rows(0)("tmpEnd_Time")
                                        If CDate(_DyeEnd_Date) > txtDate.Text Then

                                        Else
                                            _Status = True
                                            dg3.Rows(i).Cells(2).Value = _No_Of_Batch
                                            dg3.Rows(i).Cells(3).Value = _Balance_Qty
                                            dg3.Rows(i).Cells(4).Value = CInt(_Balance_Qty / 1.05)
                                            dg3.Rows(i).Cells(5).Value = M01.Tables(0).Rows(X)("M36MC_No")
                                            UltraGrid1.Rows(i).Cells(7).Value = UltraGrid1.Rows(i).Cells(7).Value + (5 * _No_Of_Batch)

                                            ncQryType = "ADD1"
                                            nvcFieldList1 = "(tmpRefNo," & "tmpMC_No," & "tmp15Class," & "tmpDate," & "tmpSTTime," & "tmpEnd_Time," & "tmpQty," & "tmpStatus," & "tmpShade," & "tmpCatagary," & "tmpCon) " & "values('" & Delivary_Ref & "','" & Trim(M01.Tables(0).Rows(X)("M36MC_No")) & "','" & dg1.Rows(i).Cells(0).Value & "','" & CDate(_DyeEnd_Date) & "','" & _DyeStartDate & "','" & _DyeEnd_Date & "','" & _Balance_Qty & "','I','" & dg1.Rows(i).Cells(2).Value & "','" & dg1.Rows(i).Cells(3).Value & "','" & CInt(_Balance_Qty / 1.05) & "')"
                                            up_GetSetYarn_DyeMCPln(ncQryType, nvcFieldList1, vcWhere, connection, transaction)
                                            'transaction.Commit()
                                            Exit For
                                        End If
                                    End If
                                Else
                                    _Status = True
                                    _DyeStartDate = _DyeStartDate & " 7:30AM"
                                    _No_Of_Batch = _No_Of_Batch + (_Balance_Qty / M01.Tables(0).Rows(X)("m36max_qty"))
                                    _DyeEnd_Date = _DyeStartDate.AddHours(+10)

                                    dg3.Rows(i).Cells(2).Value = _No_Of_Batch
                                    dg3.Rows(i).Cells(3).Value = _Balance_Qty
                                    dg3.Rows(i).Cells(4).Value = CInt(_Balance_Qty / 1.05)
                                    dg3.Rows(i).Cells(5).Value = M01.Tables(0).Rows(X)("M36MC_No")
                                    If _Machine_No <> "" Then
                                        _Machine_No = _Machine_No & "/" & M01.Tables(0).Rows(X)("M36MC_No")
                                    Else
                                        _Machine_No = M01.Tables(0).Rows(X)("M36MC_No")
                                    End If
                                    dg1.Rows(i).Cells(7).Value = dg1.Rows(i).Cells(7).Value + (5 * _No_Of_Batch)

                                    ncQryType = "ADD1"
                                    nvcFieldList1 = "(tmpRefNo," & "tmpMC_No," & "tmp15Class," & "tmpDate," & "tmpSTTime," & "tmpEnd_Time," & "tmpQty," & "tmpStatus," & "tmpShade," & "tmpCatagary," & "tmpCon) " & "values('" & Delivary_Ref & "','" & Trim(M01.Tables(0).Rows(X)("M36MC_No")) & "','" & dg1.Rows(i).Cells(0).Value & "','" & CDate(_DyeEnd_Date) & "','" & _DyeStartDate & "','" & _DyeEnd_Date & "','" & _Balance_Qty & "','I','" & dg1.Rows(i).Cells(1).Value & "','" & dg1.Rows(i).Cells(2).Value & "','" & CInt(_Balance_Qty / 1.05) & "')"
                                    up_GetSetYarn_DyeMCPln(ncQryType, nvcFieldList1, vcWhere, connection, transaction)
                                    ' transaction.Commit()
                                    Exit For
                                End If
                            End If
                            X = X + 1
                        Next
                    End If
                End If
                i = i + 1
            Next

            transaction.Commit()
            connection.Close()
            UltraTabControl1.Tabs(3).Enabled = True
            UltraTabControl1.SelectedTab = UltraTabControl1.Tabs(3)
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
                connection.Close()
            End If
        End Try
    End Sub

    Private Sub dg1_YB_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles dg1_YB.AfterCellUpdate
        Dim _Rowindex As Integer
        Dim I As Integer
        _Rowindex = dg1_YB.ActiveRow.Index
        Dim _Status As String
        Dim _Cat As String
        _Status = ""
        If chkCh1.Checked = True Then
            _Status = "1"
        ElseIf chkCh2.Checked = True Then
            _Status = "2"
        ElseIf chkCh3.Checked = True Then
            _Status = "3"
        End If
        If (Trim(dg1_YB.Rows(_Rowindex).Cells(7).Text)) = True Then

            If Microsoft.VisualBasic.Left(UCase(Trim(txtYarn1.Text)), 7) = Microsoft.VisualBasic.Left(UCase(Trim(dg1_YB.Rows(_Rowindex).Cells(1).Value)), 7) Then
                _Cat = Microsoft.VisualBasic.Left(UCase(Trim(dg1_YB.Rows(_Rowindex).Cells(1).Value)), 7)
                If chkCh1.Checked = True Then
                    '  MsgBox(dg1_YB.Rows(_Rowindex).Cells(5).Value)
                    If lblBalance_YB.Text >= CDbl(dg1_YB.Rows(_Rowindex).Cells(5).Value) Then
                        dg1_YB.Rows(_Rowindex).Cells(6).Value = dg1_YB.Rows(_Rowindex).Cells(5).Value
                        Call Update_tmp_Yarnbooking(_Rowindex, _Status, _Cat)
                        Call Calculation_YB_Balance(_Status, _Cat)
                    Else
                        Dim windowInfo As New Infragistics.Win.Misc.UltraDesktopAlertShowWindowInfo
                        Dim strFileName As String
                        windowInfo.Caption = "Balance Quantity less than Allocated qty."
                        windowInfo.FooterText = "Technova"
                        strFileName = ConfigurationManager.AppSettings("SoundPath") + "\REMINDER.wav"
                        windowInfo.Sound = strFileName
                        UltraDesktopAlert1.Show(windowInfo)
                    End If
                End If
            ElseIf Microsoft.VisualBasic.Left(UCase(Trim(txtYarn2.Text)), 7) = Microsoft.VisualBasic.Left(UCase(Trim(dg1_YB.Rows(_Rowindex).Cells(1).Value)), 7) Then
                _Cat = Microsoft.VisualBasic.Left(UCase(Trim(dg1_YB.Rows(_Rowindex).Cells(1).Value)), 7)
                If chkCh2.Checked = True Then
                    '  MsgBox(dg1_YB.Rows(_Rowindex).Cells(5).Value)
                    If lblBalance_YB1.Text >= CDbl(dg1_YB.Rows(_Rowindex).Cells(5).Value) Then
                        dg1_YB.Rows(_Rowindex).Cells(6).Value = dg1_YB.Rows(_Rowindex).Cells(5).Value
                        Call Update_tmp_Yarnbooking(_Rowindex, _Status, _Cat)
                        Call Calculation_YB_Balance(_Status, _Cat)
                    Else
                        Dim windowInfo As New Infragistics.Win.Misc.UltraDesktopAlertShowWindowInfo
                        Dim strFileName As String
                        windowInfo.Caption = "Balance Quantity less than Allocated qty."
                        windowInfo.FooterText = "Technova"
                        strFileName = ConfigurationManager.AppSettings("SoundPath") + "\REMINDER.wav"
                        windowInfo.Sound = strFileName
                        UltraDesktopAlert1.Show(windowInfo)
                    End If
                End If
            ElseIf Microsoft.VisualBasic.Left(UCase(Trim(txtYarn3.Text)), 7) = Microsoft.VisualBasic.Left(UCase(Trim(dg1_YB.Rows(_Rowindex).Cells(1).Value)), 7) Then
                _Cat = Microsoft.VisualBasic.Left(UCase(Trim(dg1_YB.Rows(_Rowindex).Cells(1).Value)), 7)
                If chkCh3.Checked = True Then
                    '  MsgBox(dg1_YB.Rows(_Rowindex).Cells(5).Value)
                    If lblBalance_YB2.Text >= CDbl(dg1_YB.Rows(_Rowindex).Cells(5).Value) Then
                        dg1_YB.Rows(_Rowindex).Cells(6).Value = dg1_YB.Rows(_Rowindex).Cells(5).Value
                        Call Update_tmp_Yarnbooking(_Rowindex, _Status, _Cat)
                        Call Calculation_YB_Balance(_Status, _Cat)
                    End If
                End If
            End If
        Else
            If Microsoft.VisualBasic.Left(UCase(Trim(txtYarn1.Text)), 7) = Microsoft.VisualBasic.Left(UCase(Trim(dg1_YB.Rows(_Rowindex).Cells(1).Value)), 7) Then
                _Cat = Microsoft.VisualBasic.Left(UCase(Trim(dg1_YB.Rows(_Rowindex).Cells(1).Value)), 7)
                If chkCh1.Checked = True Then
                    '  MsgBox(dg1_YB.Rows(_Rowindex).Cells(5).Value)
                   
                    Call Update_tmp_Yarnbooking(_Rowindex, _Status, _Cat)
                    Call Calculation_YB_Balance(_Status, _Cat)

                End If
            ElseIf Microsoft.VisualBasic.Left(UCase(Trim(txtYarn2.Text)), 7) = Microsoft.VisualBasic.Left(UCase(Trim(dg1_YB.Rows(_Rowindex).Cells(1).Value)), 7) Then
                _Cat = Microsoft.VisualBasic.Left(UCase(Trim(dg1_YB.Rows(_Rowindex).Cells(1).Value)), 7)
                If chkCh2.Checked = True Then
                    'MsgBox(dg1_YB.Rows(_Rowindex).Cells(5).Value)

                    If CDbl(dg1_YB.Rows(_Rowindex).Cells(5).Value) >= lblBalance_YB1.Text Then
                        '  dg1_YB.Rows(_Rowindex).Cells(6).Value = dg1_YB.Rows(_Rowindex).Cells(5).Value
                        Call Update_tmp_Yarnbooking(_Rowindex, _Status, _Cat)
                        Call Calculation_YB_Balance(_Status, _Cat)
                    Else
                        Dim windowInfo As New Infragistics.Win.Misc.UltraDesktopAlertShowWindowInfo
                        Dim strFileName As String
                        windowInfo.Caption = "Balance Quantity less than Allocated qty."
                        windowInfo.FooterText = "Technova"
                        strFileName = ConfigurationManager.AppSettings("SoundPath") + "\REMINDER.wav"
                        windowInfo.Sound = strFileName
                        UltraDesktopAlert1.Show(windowInfo)
                    End If
                End If
            ElseIf Microsoft.VisualBasic.Left(UCase(Trim(txtYarn3.Text)), 7) = Microsoft.VisualBasic.Left(UCase(Trim(dg1_YB.Rows(_Rowindex).Cells(1).Value)), 7) Then
                _Cat = Microsoft.VisualBasic.Left(UCase(Trim(dg1_YB.Rows(_Rowindex).Cells(1).Value)), 7)
                If chkCh3.Checked = True Then
                    '  MsgBox(dg1_YB.Rows(_Rowindex).Cells(5).Value)
                    ' If lblBalance_YB2.Text >= CDbl(dg1_YB.Rows(_Rowindex).Cells(5).Value) Then

                    dg1_YB.Rows(_Rowindex).Cells(6).Value = dg1_YB.Rows(_Rowindex).Cells(5).Value
                    Call Update_tmp_Yarnbooking(_Rowindex, _Status, _Cat)
                    Call Calculation_YB_Balance(_Status, _Cat)
                    'End If
                End If
        End If
        End If
    End Sub

    Function Calculation_YB_Balance(ByVal _Status As String, ByVal _Cat As String)
        Dim M01 As DataSet
        Dim vcWhere As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim Value As Double
        Try
            vcWhere = " tmpStatus='" & _Status & "' and tmpCat='" & _Cat & "' and tmpSO='" & strSales_Order & "' and tmpLine_Item=" & strLine_Item & ""
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TYBC"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                Value = M01.Tables(0).Rows(0)("Qty")
                If Trim(M01.Tables(0).Rows(0)("tmpStatus")) = "1" Then
                    Value = txtReq1.Text - Value
                    lblBalance_YB.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    lblBalance_YB.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                ElseIf Trim(M01.Tables(0).Rows(0)("tmpStatus")) = "2" Then
                    Value = txtReq2.Text - Value
                    lblBalance_YB1.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    lblBalance_YB1.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                ElseIf Trim(M01.Tables(0).Rows(0)("tmpStatus")) = "3" Then
                    Value = txtReq3.Text - Value
                    lblBalance_YB2.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    lblBalance_YB2.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                End If
            End If

            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            con.Close()
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)

                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
                con.Close()
            End If
        End Try
    End Function

    Function Update_tmp_Yarnbooking(ByVal _RowIndex As Integer, ByVal _Status As String, ByVal _category As String)
        Dim ncQryType As String
        Dim nvcFieldList1 As String
        Dim M02 As DataSet
        Dim vcWhere As String
        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean
        Dim M01 As DataSet
        Try
            connection = DBEngin.GetConnection(True)
            connectionCreated = True
            transaction = connection.BeginTransaction()
            transactionCreated = True
            If _Status <> "" Then
                vcWhere = "tmp10Class='" & dg1_YB.Rows(_RowIndex).Cells(0).Value & "' and tmpStock_Code='" & dg1_YB.Rows(_RowIndex).Cells(2).Value & "' and tmpLocation='" & dg1_YB.Rows(_RowIndex).Cells(4).Value & "' and tmpSO='" & strSales_Order & "' and tmpLine_Item=" & strLine_Item & ""
                M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TYBK"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    If IsNumeric(CDbl(dg1_YB.Rows(_RowIndex).Cells(6).Text)) Then
                        If CDbl(dg1_YB.Rows(_RowIndex).Cells(6).Text) > CDbl(dg1_YB.Rows(_RowIndex).Cells(5).Text) Then
                            Dim windowInfo As New Infragistics.Win.Misc.UltraDesktopAlertShowWindowInfo
                            Dim strFileName As String
                            windowInfo.Caption = "Please check the request yarn qty"
                            windowInfo.FooterText = "Technova"
                            strFileName = ConfigurationManager.AppSettings("SoundPath") + "\REMINDER.wav"
                            windowInfo.Sound = strFileName
                            UltraDesktopAlert1.Show(windowInfo)

                            DBEngin.CloseConnection(connection)
                            connection.ConnectionString = ""
                            connection.Close()
                            Exit Function
                        End If
                    End If
                    If Trim(dg1_YB.Rows(_RowIndex).Cells(6).Text) <> "" Then
                        nvcFieldList1 = "update tmpYarn_Booking set tmpQty='" & CDbl(dg1_YB.Rows(_RowIndex).Cells(6).Value) & "' where  tmp10Class='" & dg1_YB.Rows(_RowIndex).Cells(0).Value & "' and tmpStock_Code='" & dg1_YB.Rows(_RowIndex).Cells(2).Value & "' and tmpLocation='" & dg1_YB.Rows(_RowIndex).Cells(4).Value & "' and tmpSO='" & strSales_Order & "' and tmpLine_Item=" & strLine_Item & ""
                        ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                    Else
                        nvcFieldList1 = "update tmpYarn_Booking set tmpQty='0' where  tmp10Class='" & dg1_YB.Rows(_RowIndex).Cells(0).Value & "' and tmpStock_Code='" & dg1_YB.Rows(_RowIndex).Cells(2).Value & "' and tmpLocation='" & dg1_YB.Rows(_RowIndex).Cells(4).Value & "' and tmpSO='" & strSales_Order & "' and tmpLine_Item=" & strLine_Item & ""
                        ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                    End If
                Else
                    If Trim(dg1_YB.Rows(_RowIndex).Cells(6).Text) <> "" Then
                        If CDbl(dg1_YB.Rows(_RowIndex).Cells(6).Text) > CDbl(dg1_YB.Rows(_RowIndex).Cells(5).Text) Then
                            Dim windowInfo As New Infragistics.Win.Misc.UltraDesktopAlertShowWindowInfo
                            Dim strFileName As String
                            windowInfo.Caption = "Please check the request yarn qty"
                            windowInfo.FooterText = "Technova"
                            strFileName = ConfigurationManager.AppSettings("SoundPath") + "\REMINDER.wav"
                            windowInfo.Sound = strFileName
                            UltraDesktopAlert1.Show(windowInfo)

                            DBEngin.CloseConnection(connection)
                            connection.ConnectionString = ""
                            connection.Close()
                            Exit Function
                        End If
                        ncQryType = "YADD"
                        nvcFieldList1 = "(tmpRef," & "tmp10Class," & "tmpStock_Code," & "tmpLocation," & "tmpQty," & "tmpStatus," & "tmpCat," & "tmpSO," & "tmpLine_Item," & "tmpTime," & "tmpB_Status) " & "values(" & Delivary_Ref & ",'" & dg1_YB.Rows(_RowIndex).Cells(0).Value & "','" & dg1_YB.Rows(_RowIndex).Cells(2).Value & "','" & dg1_YB.Rows(_RowIndex).Cells(4).Value & "','" & CDbl(dg1_YB.Rows(_RowIndex).Cells(6).Value) & "','" & _Status & "','" & _category & "','" & strSales_Order & "'," & strLine_Item & ",'" & Now & "','A')"
                        up_GetSetYarn_Bookingtmp(ncQryType, nvcFieldList1, vcWhere, connection, transaction)
                    End If
                End If
            Else
                Dim windowInfo As New Infragistics.Win.Misc.UltraDesktopAlertShowWindowInfo
                Dim strFileName As String
                windowInfo.Caption = "Please select the Yarn"
                windowInfo.FooterText = "Technova"
                strFileName = ConfigurationManager.AppSettings("SoundPath") + "\REMINDER.wav"
                windowInfo.Sound = strFileName
                UltraDesktopAlert1.Show(windowInfo)
            End If
            transaction.Commit()

            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
            connection.Close()
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)

                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
                connection.Close()
            End If
        End Try
    End Function

    Private Sub dg1_YB_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles dg1_YB.AfterRowUpdate
        On Error Resume Next
        Dim _Rowindex As Integer
        Dim I As Integer
        _Rowindex = dg1_YB.ActiveRow.Index
        Dim _Status As String
        Dim _Cat As String

        If chkCh1.Checked = True Then
            _Status = "1"
        ElseIf chkCh2.Checked = True Then
            _Status = "2"
        ElseIf chkCh3.Checked = True Then
            _Status = "3"
        End If
        _Cat = Microsoft.VisualBasic.Left(UCase(Trim(dg1_YB.Rows(_Rowindex).Cells(1).Value)), 7)
        Call Update_tmp_Yarnbooking(_Rowindex, _Status, _Cat)
        Call Calculation_YB_Balance(_Status, _Cat)
    End Sub


    Function Load_YarnCombo()
        Dim i As Integer
        Dim _Balance As Double
        Dim dt As DataTable = New DataTable()
        ' dt.Columns.Add("ID", GetType(Integer))
        dt.Columns.Add("##", GetType(String))
       


        If lblBalance_YB.Text <> "" Then
            If (lblBalance_YB.Text) > 0 Then
                dt.Rows.Add(New Object() {Trim(txtYarn1.Text)})
                dt.AcceptChanges()
            End If
        End If

        If lblBalance_YB1.Text <> "" Then
            If (lblBalance_YB1.Text) > 0 Then
                dt.Rows.Add(New Object() {Trim(txtYarn2.Text)})
                dt.AcceptChanges()
            End If
        End If

        If lblBalance_YB2.Text <> "" Then
            If (lblBalance_YB2.Text) > 0 Then
                dt.Rows.Add(New Object() {Trim(txtYarn3.Text)})
                dt.AcceptChanges()
            End If
        End If
        ' Next
        Me.cboY_Name.SetDataBinding(dt, Nothing)
        cboY_Name.DisplayMember = "##"
        cboY_Name.Rows.Band.Columns(0).Width = 350
        '  Me.UltraDropDown1.ValueMember = "ID"

    End Function
 
    Private Sub cmdKnt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdKnt.Click
        Dim vcWhere As String
        Dim M01 As DataSet
        Dim vcFieldList As String
        Dim ncQryType As String
        Dim nvcFieldList1 As String
        Dim M02 As DataSet

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Dim i As Integer
        Try
            Call Load_YarnCombo()
            Call Load_Gride_yarn_Request()

            If OPR8.Visible = True Then
                OPR8.Visible = False
                Exit Sub
            End If
            If lblBalance_YB.Text <> "" Then
                If (lblBalance_YB.Text) > 0 Then
                    OPR8.Visible = True

                    Exit Sub
                End If
            End If
            If lblBalance_YB1.Text <> "" Then
                If lblBalance_YB1.Text > 0 Then
                    OPR8.Visible = True
                    Exit Sub
                End If
            End If
            If lblBalance_YB2.Text <> "" Then
                If lblBalance_YB2.Text > 0 Then
                    OPR8.Visible = True
                    Exit Sub
                End If
            End If

            If OPR8.Visible = True Then
                lblShade.Visible = False
            Else
                lblShade.Visible = True
            End If
            strKnitting_PlanStatus = "Yarn Booking"
            If txtReq1.Text <> "" And txtReq2.Text <> "" And txtReq3.Text <> "" Then
                strQty = CDbl(txtReq1.Text) + CDbl(txtReq2.Text) + CDbl(txtReq3.Text)
            ElseIf txtReq1.Text <> "" And txtReq2.Text <> "" Then
                strQty = CDbl(txtReq1.Text) + CDbl(txtReq2.Text)
            ElseIf txtReq1.Text <> "" Then
                strQty = CDbl(txtReq1.Text)
            End If

            i = 0
            Dim _status As Boolean
            M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "BKD"))
            If isValidDataset(M01) Then
                If Trim(M01.Tables(0).Rows(0)("tmpUser")) = strDisname Then
                    If lblBalance_YB.Text <> "" Then
                    Else
                        lblBalance_YB.Text = "0.0"
                    End If
                    If CDbl(lblBalance_YB.Text) = 0 Then
                        UltraTabControl1.Tabs(4).Enabled = True
                        UltraTabControl1.SelectedTab = UltraTabControl1.Tabs(4)
                    Else
                        '1st Yarn Checking
                        For Each uRow As UltraGridRow In dg1_YB.Rows
                            If Microsoft.VisualBasic.Left(Trim(dg1_YB.Rows(i).Cells(1).Text), 4) = Microsoft.VisualBasic.Left(txtYarn1.Text, 4) Then
                                _status = True
                                Exit For
                            End If
                            i = i + 1
                        Next

                        If _status = False Then
                            MsgBox("Please request the Yarn", MsgBoxStyle.Information, "Information .....")
                            OPR8.Visible = True
                            Exit Sub
                        End If

                        _status = False
                        i = 0
                        For Each uRow As UltraGridRow In dg1_YB.Rows
                            If Microsoft.VisualBasic.Left(Trim(dg1_YB.Rows(i).Cells(1).Text), 4) = Microsoft.VisualBasic.Left(txtYarn2.Text, 4) Then
                                _status = True
                                Exit For
                            End If
                            i = i + 1
                        Next


                        If _status = False Then
                            MsgBox("Please request the Yarn", MsgBoxStyle.Information, "Information .....")
                            OPR8.Visible = True
                            Exit Sub
                        End If

                        MsgBox("Can't Plan Knitting Machine yet", MsgBoxStyle.Information, "Information .......")
                        Exit Sub
                    End If

                Else
                    MsgBox("System used by " & M01.Tables(0).Rows(0)("tmpUser") & " Please try again", MsgBoxStyle.Information, "Information ......")
                    connection.Close()
                    Exit Sub
                End If
            Else
                If lblBalance_YB.Text <> "" Then
                Else
                    lblBalance_YB.Text = "0.00"
                End If
                If CDbl(lblBalance_YB.Text) = 0 Then
                    ncQryType = "BADD"
                    nvcFieldList1 = "(tmpSales_Order," & "tmpLine_Item," & "tmpUser) " & "values('" & strSales_Order & "'," & strLine_Item & ",'" & strDisname & "')"
                    up_GetSetBlock_KntMC(ncQryType, nvcFieldList1, vcWhere, connection, transaction)

                    transaction.Commit()
                    connection.Close()
                    UltraTabControl1.Tabs(4).Enabled = True
                    UltraTabControl1.SelectedTab = UltraTabControl1.Tabs(4)
                Else
                    MsgBox("Can't Plan Knitting Machine yet", MsgBoxStyle.Information, "Information .......")
                    connection.Close()
                    Exit Sub
                End If

            End If

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)

                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
                connection.Close()
            End If
        End Try
    End Sub

    Function Update_Records()
        Dim vcWhere As String
        Dim M01 As DataSet
        Dim vcFieldList As String
        Dim ncQryType As String
        Dim nvcFieldList1 As String
        Dim M02 As DataSet

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Dim i As Integer
        Dim _Balance As Double
        i = 0
        '1st Yarn
        Try
            If txtYarn1.Text <> "" Then
                For Each uRow As UltraGridRow In UltraGrid1.Rows
                    If Microsoft.VisualBasic.Left(Trim(UltraGrid1.Rows(i).Cells(1).Text), 4) = Microsoft.VisualBasic.Left(txtYarn1.Text, 4) Then
                        If Trim(UltraGrid1.Rows(i).Cells(6).Text) <> "" Then
                            _Balance = _Balance + Trim(UltraGrid1.Rows(i).Cells(6).Value)
                        End If
                    End If

                    i = i + 1
                Next
            End If

            If (CDbl(txtReq1.Text) - _Balance) > 0 Then
                vcWhere = "T14Ref_no=" & Delivary_Ref & " and T14Sales_order='" & strSales_Order & "' and T14Line_Item=" & strLine_Item & " and T14Yarn='" & txtYarn1.Text & "'"
                M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "YRQ"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then

                Else
                    ncQryType = "AYR"
                    nvcFieldList1 = "(T14Ref_no," & "T14Sales_order," & "T14Line_Item," & "T14Yarn," & "T14Req_By," & "T14Req_Date," & "T14Status," & "T14Qty," & "T14Time) " & "values(" & Delivary_Ref & ",'" & strSales_Order & "'," & strLine_Item & ",'" & txtYarn1.Text & "','" & strDisname & "','" & Today & "','N','" & CDbl(txtReq1.Text) - _Balance & "','" & Now & "')"
                    up_GetSetYarn_Request(ncQryType, nvcFieldList1, vcWhere, connection, transaction)


                End If
            End If
            '2nd Yarn
            i = 0
            _Balance = 0
            If txtYarn2.Text <> "" Then
                For Each uRow As UltraGridRow In dg1_YB.Rows
                    If Microsoft.VisualBasic.Left(Trim(dg1_YB.Rows(i).Cells(1).Text), 4) = Microsoft.VisualBasic.Left(txtYarn2.Text, 4) Then
                        If Trim(dg1_YB.Rows(i).Cells(6).Text) <> "" Then
                            _Balance = _Balance + Trim(dg1_YB.Rows(i).Cells(6).Value)
                        End If
                    End If

                    i = i + 1
                Next
            End If

            If (CDbl(txtReq2.Text) - _Balance) > 0 Then
                vcWhere = "T14Ref_no=" & Delivary_Ref & " and T14Sales_order='" & strSales_Order & "' and T14Line_Item=" & strLine_Item & " and T14Yarn='" & txtYarn2.Text & "'"
                M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "YRQ"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then

                Else
                    ncQryType = "AYR"
                    nvcFieldList1 = "(T14Ref_no," & "T14Sales_order," & "T14Line_Item," & "T14Yarn," & "T14Req_By," & "T14Req_Date," & "T14Status," & "T14Qty," & "T14Time) " & "values(" & Delivary_Ref & ",'" & strSales_Order & "'," & strLine_Item & ",'" & txtYarn2.Text & "','" & strDisname & "','" & Today & "','N','" & CDbl(txtReq2.Text) - _Balance & "','" & Now & "')"
                    up_GetSetYarn_Request(ncQryType, nvcFieldList1, vcWhere, connection, transaction)


                End If
            End If

            transaction.Commit()
            connection.Close()
            Me.Close()

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)

                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
                connection.Close()
            End If
        End Try
    End Function

    Private Sub UltraButton11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton11.Click
        Me.Close()
    End Sub

    Private Sub UltraButton10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton10.Click
        frmKnitting_Plan_Board.Show()
    End Sub

  
    Private Sub UltraButton16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton16.Click
        Me.Close()
    End Sub

    Private Sub UltraButton14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton14.Click
        dg2_YB.Visible = False
        With UltraGroupBox24
            .Location = New Point(25, 100)
            .Width = 806
            .Height = 295

        End With
        With dg1_YB
            .Width = 794
            .Height = 247
        End With

    End Sub

    Function Load_Gride_Knt()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1_KNT = CustomerDataClass.MakeDataTableKnitting_PLN
        UltraGrid4.DataSource = c_dataCustomer1_KNT
        With UltraGrid4
            .DisplayLayout.Bands(0).Columns(0).Width = 70
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 110
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 110
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(3).Width = 90
            .DisplayLayout.Bands(0).Columns(3).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(4).Width = 70
            .DisplayLayout.Bands(0).Columns(4).AutoEdit = False

            .DisplayLayout.Bands(0).Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Left
            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Left
            .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

            '.DisplayLayout.Bands(0).Columns(3).Width = 90
            '.DisplayLayout.Bands(0).Columns(4).Width = 90
            '.DisplayLayout.Bands(0).Columns(5).Width = 90
            ''  .DisplayLayout.Bands(0).Columns(6).Width = 90
            ' .DisplayLayout.Bands(0).Columns(7).Width = 90

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function

    Function Load_Gride_Delivary()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1_Delivary = CustomerDataClass.MakeDataTable_Delivary
        UltraGrid3.DataSource = c_dataCustomer1_Delivary
        With UltraGrid3
            .DisplayLayout.Bands(0).Columns(0).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 110
            '.DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            '.DisplayLayout.Bands(0).Columns(2).Width = 110
            '.DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            '.DisplayLayout.Bands(0).Columns(3).Width = 90
            '.DisplayLayout.Bands(0).Columns(3).AutoEdit = False
            '.DisplayLayout.Bands(0).Columns(4).Width = 90
            '.DisplayLayout.Bands(0).Columns(4).AutoEdit = False

            .DisplayLayout.Bands(0).Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Left
            '.DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            '.DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center


            '.DisplayLayout.Bands(0).Columns(3).Width = 90
            '.DisplayLayout.Bands(0).Columns(4).Width = 90
            '.DisplayLayout.Bands(0).Columns(5).Width = 90
            ''  .DisplayLayout.Bands(0).Columns(6).Width = 90
            ' .DisplayLayout.Bands(0).Columns(7).Width = 90

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function

    Function Load_Gride_Delivary_Daily()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1_Delivary = CustomerDataClass.MakeDataTable_Delivary_Daily
        UltraGrid3.DataSource = c_dataCustomer1_Delivary
        With UltraGrid3
            .DisplayLayout.Bands(0).Columns(0).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 110
            '.DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            '.DisplayLayout.Bands(0).Columns(2).Width = 110
            '.DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            '.DisplayLayout.Bands(0).Columns(3).Width = 90
            '.DisplayLayout.Bands(0).Columns(3).AutoEdit = False
            '.DisplayLayout.Bands(0).Columns(4).Width = 90
            '.DisplayLayout.Bands(0).Columns(4).AutoEdit = False

            .DisplayLayout.Bands(0).Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
           
        End With
    End Function

    Private Sub UltraButton13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton13.Click
        dg2_YB.Visible = True
        With UltraGroupBox24
            .Location = New Point(25, 218)
            .Width = 806
            .Height = 190

        End With
        With dg1_YB
            .Width = 794
            .Height = 137
        End With
    End Sub

    Private Sub UltraButton18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton18.Click
        If lblKnt_Balance.Text > 0 Then
            
        Else
            Exit Sub
        End If

        If IsNumeric(txtK_Qty.Text) Then
            If CDbl(txtK_Qty.Text) > CDbl(lblKnt_Balance.Text) Then
                MsgBox("Please enter the correct Qty", MsgBoxStyle.Information, "Information ....")
                Exit Sub
            End If
        Else
            MsgBox("Please enter the correct Qty", MsgBoxStyle.Information, "Information ....")
            Exit Sub
        End If
        If CDbl(lblKP_Balance.Text) >= 0 Then
            lblKP_Balance.Text = txtK_Qty.Text
        End If
        Call Load_Gride_projection()
        '  Call Load_Projection_Detailes(txtDelivary_Date.Text)
        Call Save_Data(Today)
        Call Calculate_Tobe_Knt_Balance()
    End Sub
    Function Save_Data(ByVal strDate As Date)
        Dim vcWhere As String
        Dim M01 As DataSet
        Dim vcFieldList As String
        Dim ncQryType As String
        Dim nvcFieldList1 As String
        Dim M02 As DataSet

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Try
            Call Load_Gride_Knt()

            'vcWhere = "tmpRef_No=" & Delivary_Ref & ""
            'M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "DKP"), New SqlParameter("@vcWhereClause1", vcWhere))

            'transaction.Commit()
            'connection.Close()
            'If txtComplete_Date_Knt.Text <= strDate Then
            '    MsgBox("Can't Plan this day.Please try again", MsgBoxStyle.Information, "Information .....")
            '    Exit Function
            'End If

            If IsNumeric(txtK_Qty.Text) Then
            Else
                MsgBox("Please enter the correct Qty", MsgBoxStyle.Information, "Information .....")
                txtK_Qty.Focus()
                Exit Function
            End If

            If txtK_Qty.Text <> "" Then
            Else
                MsgBox("Please enter the Qty", MsgBoxStyle.Information, "Information .....")
                txtK_Qty.Focus()
                Exit Function
            End If
            nvcFieldList1 = "Update T15Projection_Allocation set T15Status='Y' where T15Sales_Order='" & strSales_Order & "' and T15Line_Item=" & strLine_Item & ""
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            nvcFieldList1 = "delete from tmp_Knt_Mc where tmpSales_Order='" & strSales_Order & "' and tmpLine_Item=" & strLine_Item & ""
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            nvcFieldList1 = "delete from tmpBlock_KntMC where tmpUser='" & strDisname & "' "
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            transaction.Commit()
            Call Search_Available_KMCNew(strDate)
           
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)

                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
                connection.Close()
            End If
        End Try
    End Function

    Function Update_Tempary_Knt_Mc(ByVal strMC As String, ByVal strQtype As String)
        Dim vcWhere As String
        Dim M01 As DataSet
        Dim vcFieldList As String
        Dim ncQryType As String
        Dim nvcFieldList1 As String
        Dim M02 As DataSet

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Try
            'vcWhere = "tmpSales_Order='" & strSales_Order & "' and tmpLine_Item='" & strLine_Item & "' and tmpMC_No='" & strMC & "'"
            'M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSettmp_Knt_Mc", New SqlParameter("@cQryType", strQtype), New SqlParameter("@vcWhereClause1", vcWhere))
            'If isValidDataset(M01) Then

            'Else
            ncQryType = strQtype
            nvcFieldList1 = "(tmpSales_Order," & "tmpLine_Item," & "tmpMC_No) " & "values('" & strSales_Order & "'," & strLine_Item & ",'" & strMC & "')"
            up_GetSettmp_Knt_Mc(ncQryType, nvcFieldList1, vcWhere, connection, transaction)
            ' End If
            transaction.Commit()
            connection.Close()
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)

                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
                connection.Close()
            End If
        End Try
    End Function
    Function Calculate_Tobe_Knt_Balance()
        Dim M01 As DataSet
        Dim vcWhere As String
        Dim Value As Double
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Try
            vcWhere = "tmpSales_Order='" & strSales_Order & "' and tmpLine_Item=" & strLine_Item & ""
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "SBQ"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                Value = txtQty_Knt.Text - M01.Tables(0).Rows(0)("QTY")
                lblKnt_Balance.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                lblKnt_Balance.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            End If
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            con.close()

            'Dim dt = New DataTable
            'dt.Columns.Add("Name", GetType(String))
            'Dim r = dt.NewRow()
            'r("Name") = "Set1"
            'dt.Rows.Add(r)
            'r = dt.NewRow
            'r("Name") = "Set2"
            'dt.Rows.Add(r)
            'cboK_Mc.DataSource = dt

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)

                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
                con.close()
            End If
        End Try
    End Function

    Private Sub txtComplete_Date_Knt_AfterDropDown(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtComplete_Date_Knt.AfterDropDown
        Call Search_WeekNo()
    End Sub

    Function Search_Available_KMCNew(ByVal strDate As Date)
        Dim vcWhere As String
        Dim M01 As DataSet
        Dim vcFieldList As String
        Dim ncQryType As String
        Dim nvcFieldList1 As String
        Dim M02 As DataSet
        Dim T04 As DataSet
        Dim _DAYKNIT As Integer
        Dim Value As Double
        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean
        Dim strQty As String

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Dim _Balance_Qty As Double
        Dim _MinQty As Double
        Dim i As Integer
        Dim _FromDate As Date
        Dim _Todate As Date
        Dim _TimeSpam As TimeSpan
        Dim _TotalTime As Double
        Dim _AllocateMC As Integer
        Dim _Knited_Time As Integer
        Dim _StartDate As Date
        Dim _EndDate As Date
        Dim _WeekNo As Integer
        Dim _UseMCNo As Integer
        Dim _McName As String
        Dim x As Integer
        Dim _Quality As String
        Dim _Qty_Minite As Double

        _TotalTime = 0
        If Microsoft.VisualBasic.Left(txtMC_Group_Knt.Text, 1) = "S" Then
            _McName = "SJ"
        ElseIf Microsoft.VisualBasic.Left(txtMC_Group_Knt.Text, 1) = "R" Then
            _McName = "DJ"

        End If
        Try
            _Qty_Minite = txtDaily_Capacity.Text / (24 * 60)
            _Balance_Qty = txtK_Qty.Text
            i = 0
            _AllocateMC = 0

            If _McName = "SJ" Then
                vcWhere = "M38MC='" & _McName & "' and M38Group='" & strGuarge & "'"
            Else
                vcWhere = "M38MC='" & _McName & "' and LEFT(M38Group,2)='" & Microsoft.VisualBasic.Left(strGuarge, 2) & "'"
            End If
            M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "MCNO"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                _UseMCNo = M01.Tables(0).Rows(0)("M38Mc_Count")
            End If

            vcWhere = "tmpQuality='" & txtQuality.Text & "' "
            M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "KPLA"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                _Quality = txtQuality.Text
            Else
                vcWhere = "tmpQuality='" & txtCommon.Text & "' "
                M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "KPLA"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    _Quality = txtCommon.Text
                Else

                End If
            End If
            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                Dim userDate As Date
                Dim StartdateofYear As Date
                ' Dim dayName As String
                Dim Dateofweek As Integer
                Dim WeekStartdate As Date
                Dim WeekEnddate As Date
                _Knited_Time = 0

                StartdateofYear = "1/1/" & txtYear_Knt.Text

                Dateofweek = (7 * txtWeek_Knt.Text) - 7
                WeekStartdate = CDate(StartdateofYear).AddDays(+Dateofweek)
                Dim thisCulture = Globalization.CultureInfo.CurrentCulture
                Dim dayOfWeek As DayOfWeek = thisCulture.Calendar.GetDayOfWeek(WeekStartdate)
                Dim dayName As String = thisCulture.DateTimeFormat.GetDayName(dayOfWeek)

                If dayName = "Sunday" Then
                    WeekStartdate = CDate(WeekStartdate).AddDays(-3)
                ElseIf dayName = "Tuesday" Then
                    WeekStartdate = CDate(WeekStartdate).AddDays(-4)
                ElseIf dayName = "Wednesday" Then
                    WeekStartdate = CDate(WeekStartdate).AddDays(-5)
                ElseIf dayName = "Thuesday" Then
                    WeekStartdate = CDate(WeekStartdate).AddDays(-3)
                ElseIf dayName = "Friday" Then
                    WeekStartdate = CDate(WeekStartdate).AddDays(-1)
                ElseIf dayName = "Saturday" Then
                    WeekStartdate = CDate(WeekStartdate).AddDays(-2)
                End If

                WeekEnddate = CDate(WeekStartdate).AddDays(+7)
                WeekStartdate = WeekStartdate & " " & "7:30AM"
                WeekEnddate = WeekEnddate & " " & "7:30AM"

                If M01.Tables(0).Rows(i)("tmpEnd_Date") >= Today Then
                    '_FromDate = CDate(txtComplete_Date_Knt.Text).AddDays(-7)
                    '_FromDate = _FromDate & " " & "7:30 AM"
                    '_Todate = txtComplete_Date_Knt.Text & " " & "7:30 AM"

                    'userDate = DateTime.Parse("1/1/" & Year(txtDate.Text))
                    '' MsgBox(WeekdayName(Weekday(userDate)))
                    'If WeekdayName(Weekday(userDate)) = "Sunday" Then
                    '    userDate = userDate.AddDays(-3)
                    'ElseIf WeekdayName(Weekday(userDate)) = "Monday" Then
                    '    userDate = userDate.AddDays(-4)
                    'ElseIf WeekdayName(Weekday(userDate)) = "Tuesday" Then
                    '    userDate = userDate.AddDays(-5)
                    'ElseIf WeekdayName(Weekday(userDate)) = "Thusday" Then
                    '    'userDate = userDate.AddDays(-1)
                    'ElseIf WeekdayName(Weekday(userDate)) = "Friday" Then
                    '    userDate = userDate.AddDays(-1)
                    'ElseIf WeekdayName(Weekday(userDate)) = "Saturday" Then
                    '    userDate = userDate.AddDays(-2)

                    'End If
                    'If IsNumeric(txtWeek_Knt.Text) Then
                    '    userDate = userDate.AddDays(txtWeek_Knt.Text * 7)
                    'Else
                    '    Exit Function
                    'End If
                Else
                    _FromDate = M01.Tables(0).Rows(i)("tmpEnd_time")
                    _Todate = txtComplete_Date_Knt.Text & " " & "7:30 AM"
                End If

                _TimeSpam = _Todate.Subtract(_FromDate)
                _TimeSpam = WeekEnddate.Subtract(WeekStartdate)
                _TotalTime = (_TimeSpam.TotalMinutes * _Qty_Minite) '/ 1000
                If _TotalTime > 0 Then
                Else
                    i = i + 1
                    Continue For
                End If
                If _TotalTime > txtK_Qty.Text Then
                    _Balance_Qty = 0
                    ' _AllocateMC = _AllocateMC + 1
                    _Knited_Time = txtK_Qty.Text / _Qty_Minite
                    If _Knited_Time > 0 Then

                        _Todate = WeekStartdate.AddMinutes(_Knited_Time)
                    End If

                    '_TimeSpam = _Todate.Subtract(_FromDate)
                    '_StartDate = (Month(_FromDate) & "/" & Microsoft.VisualBasic.Day(_FromDate) & "/" & Year(_FromDate))
                    '_EndDate = (Month(_Todate) & "/" & Microsoft.VisualBasic.Day(_Todate) & "/" & Year(_Todate))

                    '_WeekNo = DatePart("WW", _EndDate, FirstDayOfWeek.Monday)

                    'DEACTIVE BY SURANGA ON 2016.6.17
                    'ncQryType = "KADD"
                    'nvcFieldList1 = "(tmpRef_No," & "tmpMC_No," & "tmpGroup," & "tmpQuality," & "tmp20Class," & "tmpSales_Order," & "tmpLine_Item," & "tmpKnt_Order," & "tmpBalance," & "tmpWeek_No," & "tmpYear," & "tmpStart_Time," & "tmpEnd_Time," & "tmpStart_Date," & "tmpEnd_Date," & "tmpStatus," & "tmpQ_Status) " & "values(" & Delivary_Ref & ",'" & M01.Tables(0).Rows(i)("tmpMC_No") & "','" & txtMC_Group_Knt.Text & "','" & _Quality & "','" & str20Class & "','" & strSales_Order & "'," & strLine_Item & ",'" & txtK_Qty.Text & "','" & txtQty_Knt.Text & "'," & _WeekNo & "," & Year(_EndDate) & ",'" & _FromDate & "','" & _Todate & "','" & _StartDate & "','" & _EndDate & "','" & txtQuality_Group.Text & "','S')"
                    'up_GetSetBlock_KntPlanning_Boad(ncQryType, nvcFieldList1, vcWhere, connection, transaction)

                    'transaction.Commit()
                    vcWhere = "SELECT * FROM tmpKnitting_Plan_Board WHERE tmpMC_No='" & Trim(M01.Tables(0).Rows(i)("tmpMC_No")) & "' AND tmpEnd_Time>='" & _Todate & "'"
                    ' dsUser = DBEngin.ExecuteDataset(connection, transaction, "up_GetSettmp_Knt_Mc", New SqlParameter("@cQryType", "CLS1"), New SqlParameter("@vcWhereClause1", vcWhere))
                    dsUser = DBEngin.ExecuteDataset(connection, transaction, vcWhere)
                    If isValidDataset(dsUser) Then
                        i = i + 1
                        Continue For
                    Else
                        vcFieldList = "SELECT * FROM tmpKnitting_Plan_Board WHERE tmpMC_No='" & Trim(M01.Tables(0).Rows(i)("tmpMC_No")) & "' AND tmpEnd_Time BETWEEN '" & WeekStartdate & "' AND '" & WeekEnddate & "' ORDER BY tmpRef_No DESC"
                        dsUser = DBEngin.ExecuteDataset(connection, transaction, vcFieldList)
                        If isValidDataset(dsUser) Then
                            Dim newRow As DataRow = c_dataCustomer1_KNT.NewRow

                            newRow("Machine No") = M01.Tables(0).Rows(i)("tmpMC_No")
                            newRow("Start Date") = _FromDate
                            newRow("End Date") = _Todate

                            newRow("Qty") = txtK_Qty.Text
                            newRow("No of Hour") = (_TimeSpam.Days * 24 + _TimeSpam.Hours) & "." & _TimeSpam.Minutes
                            newRow("Status") = "Same Quality"
                            newRow("##") = False
                            c_dataCustomer1_KNT.Rows.Add(newRow)

                            Call Update_Tempary_Knt_Mc(M01.Tables(0).Rows(i)("tmpMC_No"), "BADD")

                        Else
                            Dim newRow As DataRow = c_dataCustomer1_KNT.NewRow
                            Dim _KNYQTY As Double

                            newRow("Machine No") = M01.Tables(0).Rows(i)("tmpMC_No")
                            newRow("Start Date") = WeekStartdate
                            If _Todate > WeekEnddate Then
                                newRow("End Date") = WeekEnddate
                                _TimeSpam = WeekEnddate.Subtract(WeekStartdate)
                                _KNYQTY = _Qty_Minite * (_TimeSpam.Days * 24 * 60)
                                _KNYQTY = _KNYQTY + (_Qty_Minite * (_TimeSpam.Hours * 60))
                                _KNYQTY = _KNYQTY + (_TimeSpam.Minutes * _Qty_Minite)
                            Else
                                newRow("End Date") = _Todate
                                _TimeSpam = _Todate.Subtract(WeekStartdate)
                                _KNYQTY = _Qty_Minite * (_TimeSpam.Days * 24 * 60)
                                _KNYQTY = _KNYQTY + (_Qty_Minite * (_TimeSpam.Hours * 60))
                                _KNYQTY = _KNYQTY + (_TimeSpam.Minutes * _Qty_Minite)
                            End If
                            newRow("Qty") = Microsoft.VisualBasic.Format(_KNYQTY, "#.00")
                            newRow("No of Hour") = (_TimeSpam.Days * 24 + _TimeSpam.Hours) & "." & _TimeSpam.Minutes
                            newRow("Status") = "Same Quality"
                            newRow("##") = False
                            c_dataCustomer1_KNT.Rows.Add(newRow)

                            Call Update_Tempary_Knt_Mc(M01.Tables(0).Rows(i)("tmpMC_No"), "BADD")
                        End If
                    End If
                        'Exit Function
                Else

                        _Balance_Qty = _Balance_Qty - _TotalTime

                        'If txtAlocate_MC.Text < _AllocateMC Then
                        '    MsgBox("No Available Machine Capacity. Please add the allocate machine", MsgBoxStyle.Information, "Information ....")
                        '    vcWhere = "tmpRef_No=" & Delivary_Ref & ""
                        '    M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "DKP"), New SqlParameter("@vcWhereClause1", vcWhere))
                        '    'transaction.Commit()
                        '    'connection.Close()
                        '    Exit Function
                        'Else

                        userDate = DateTime.Parse("1/1/" & Year(txtDate.Text))
                        ' MsgBox(WeekdayName(Weekday(userDate)))
                        If WeekdayName(Weekday(userDate)) = "Sunday" Then
                            userDate = userDate.AddDays(-3)
                        ElseIf WeekdayName(Weekday(userDate)) = "Monday" Then
                            userDate = userDate.AddDays(-4)
                        ElseIf WeekdayName(Weekday(userDate)) = "Tuesday" Then
                            userDate = userDate.AddDays(-5)
                        ElseIf WeekdayName(Weekday(userDate)) = "Thusday" Then
                            'userDate = userDate.AddDays(-1)
                        ElseIf WeekdayName(Weekday(userDate)) = "Friday" Then
                            userDate = userDate.AddDays(-1)
                        ElseIf WeekdayName(Weekday(userDate)) = "Saturday" Then
                            userDate = userDate.AddDays(-2)

                        End If
                        If IsNumeric(txtWeek_Knt.Text) Then
                            userDate = userDate.AddDays(txtWeek_Knt.Text * 7)
                        Else
                            Exit Function
                        End If

                        _StartDate = (Month(_FromDate) & "/" & Microsoft.VisualBasic.Day(_FromDate) & "/" & Year(_FromDate))
                        _EndDate = (Month(_Todate) & "/" & Microsoft.VisualBasic.Day(_Todate) & "/" & Year(_Todate))
                        _WeekNo = DatePart("WW", _EndDate, FirstDayOfWeek.Monday)


                        'DEACTIVE BY SURANGA ON 2016.6.17
                        'ncQryType = "KADD"
                        'nvcFieldList1 = "(tmpRef_No," & "tmpMC_No," & "tmpGroup," & "tmpQuality," & "tmp20Class," & "tmpSales_Order," & "tmpLine_Item," & "tmpKnt_Order," & "tmpBalance," & "tmpWeek_No," & "tmpYear," & "tmpStart_Time," & "tmpEnd_Time," & "tmpStart_Date," & "tmpEnd_Date," & "tmpStatus," & "tmpQ_Status) " & "values(" & Delivary_Ref & ",'" & M01.Tables(0).Rows(i)("tmpMC_No") & "','" & txtMC_Group.Text & "','" & _Quality & "','" & str20Class & "','" & strSales_Order & "'," & strLine_Item & ",'" & txtK_Qty.Text & "','" & _TotalTime & "'," & _WeekNo & "," & Year(_EndDate) & ",'" & _FromDate & "','" & _Todate & "','" & _StartDate & "','" & _EndDate & "','" & txtQuality_Group.Text & "','S')"
                        'up_GetSetBlock_KntPlanning_Boad(ncQryType, nvcFieldList1, vcWhere, connection, transaction)

                        ' transaction.Commit()
                        vcWhere = "tmpMC_No='" & M01.Tables(0).Rows(i)("tmpMC_No") & "' and tmpStart_Time>='" & _FromDate & "' and tmpEnd_Time<='" & _Todate & "'"
                        dsUser = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "KPLA"), New SqlParameter("@vcWhereClause1", vcWhere))
                        If isValidDataset(dsUser) Then
                            i = i + 1
                            Continue For
                        Else
                            vcWhere = "tmpMC_No='" & M01.Tables(0).Rows(i)("tmpMC_No") & "' and tmpSales_Order='" & strSales_Order & "' and tmpLine_Item='" & strLine_Item & "'"
                            dsUser = DBEngin.ExecuteDataset(connection, transaction, "up_GetSettmp_Knt_Mc", New SqlParameter("@cQryType", "CLS1"), New SqlParameter("@vcWhereClause1", vcWhere))
                            If isValidDataset(dsUser) Then
                                i = i + 1
                                Continue For
                            End If

                            Dim newRow As DataRow = c_dataCustomer1_KNT.NewRow

                            newRow("Machine No") = M01.Tables(0).Rows(i)("tmpMC_No")
                            newRow("Start Date") = _FromDate
                            newRow("End Date") = _Todate
                            Value = _TotalTime
                            strQty = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                            strQty = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                            newRow("Qty") = strQty
                            newRow("No of Hour") = (_TimeSpam.Days * 24 + _TimeSpam.Hours) & "." & _TimeSpam.Minutes
                            newRow("Status") = "Same Quality"
                            newRow("##") = False
                            c_dataCustomer1_KNT.Rows.Add(newRow)
                            _AllocateMC = _AllocateMC + 1

                            Call Update_Tempary_Knt_Mc(M01.Tables(0).Rows(i)("tmpMC_No"), "BADD")
                        End If
                End If
                    i = i + 1
            Next
            '-----------------------------------------------------------------------------------------
            'If _Balance_Qty > 0 Then

            _Quality = txtQuality.Text
            'Quality Change
            x = 0
            vcWhere = "M40Group_Name='" & Trim(txtQuality_Group.Text) & "' "
            M02 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "MGN"), New SqlParameter("@vcWhereClause1", vcWhere))
            For Each DTRow3 As DataRow In M02.Tables(0).Rows
                i = 0
                If _Quality <> "" Then
                    vcWhere = "tmpQuality<>'" & _Quality & "' and left(tmpGroup,2)='" & _McName & "' and tmpStatus='" & Trim(M02.Tables(0).Rows(x)("M40Priority_Group")) & "'"
                Else
                    vcWhere = " left(tmpGroup,2)='" & _McName & "' and tmpStatus='" & Trim(M02.Tables(0).Rows(x)("M40Priority_Group")) & "'"
                End If
                M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "KPLA"), New SqlParameter("@vcWhereClause1", vcWhere))

                If isValidDataset(M01) Then
                    For Each DTRow4 As DataRow In M01.Tables(0).Rows
                        vcWhere = "M38Group='" & strGuarge & "' and  M39Mc_No='" & Trim(M01.Tables(0).Rows(i)("tmpMC_No")) & "'"
                        dsUser = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "CAMC"), New SqlParameter("@vcWhereClause1", vcWhere))
                        If isValidDataset(dsUser) Then
                        Else
                            i = i + 1
                            Continue For
                        End If

                        vcWhere = "tmpRef_No=" & Delivary_Ref & " and  tmpMC_No='" & Trim(M01.Tables(0).Rows(i)("tmpMC_No")) & "'"
                        dsUser = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "KPCK"), New SqlParameter("@vcWhereClause1", vcWhere))
                        If isValidDataset(dsUser) Then
                            i = i + 1
                            Continue For
                        End If

                        _Knited_Time = 0
                        'If _AllocateMC > txtAlocate_MC.Text Or _Balance_Qty = 0 Then
                        '    'connection.Close()
                        '    Exit For
                        'End If
                        'MsgBox(M01.Tables(0).Rows(i)("tmpEnd_Date"))
                        Dim _Qstatus As String

                        If M01.Tables(0).Rows(i)("tmpEnd_Date") < strDate Then
                            _FromDate = strDate & " " & "7:30 AM"
                            _Todate = txtDate.Text & " " & "7:30 AM"

                        Else
                            _FromDate = M01.Tables(0).Rows(i)("tmpEnd_Time")
                            _Todate = txtDate.Text & " " & "7:30 AM"
                            'If txtQuality.Text = M01.Tables(0).Rows(i)("tmpQuality") Then
                            '    _Qstatus = "S"
                            'Else
                            _Qstatus = "QC"
                            _FromDate = _FromDate.AddHours(M02.Tables(0).Rows(x)("M40Mc_Change_HR"))
                            ' End If
                        End If

                        _TimeSpam = _Todate.Subtract(_FromDate)
                        _TotalTime = (_TimeSpam.TotalMinutes * _MinQty) '/ 1000
                        If _TotalTime > 0 Then
                        Else
                            i = i + 1
                            Continue For
                        End If
                        If _TotalTime > txtK_Qty.Text Then
                            _Balance_Qty = 0
                            _AllocateMC = _AllocateMC + 1
                            _Knited_Time = txtQty_Knt.Text / _MinQty
                            If _Knited_Time > 0 Then
                                _Todate = _FromDate.AddMinutes(+_Knited_Time)
                            End If

                            _StartDate = (Month(_FromDate) & "/" & Microsoft.VisualBasic.Day(_FromDate) & "/" & Year(_FromDate))
                            _EndDate = (Month(_Todate) & "/" & Microsoft.VisualBasic.Day(_Todate) & "/" & Year(_Todate))

                            _WeekNo = DatePart("WW", _EndDate, FirstDayOfWeek.Monday)


                            'DEACTIVE BY SURANGA ON 2016.6.17
                            'ncQryType = "KADD"
                            'nvcFieldList1 = "(tmpRef_No," & "tmpMC_No," & "tmpGroup," & "tmpQuality," & "tmp20Class," & "tmpSales_Order," & "tmpLine_Item," & "tmpKnt_Order," & "tmpBalance," & "tmpWeek_No," & "tmpYear," & "tmpStart_Time," & "tmpEnd_Time," & "tmpStart_Date," & "tmpEnd_Date," & "tmpStatus," & "tmpQ_Status) " & "values(" & Delivary_Ref & ",'" & M01.Tables(0).Rows(i)("tmpMC_No") & "','" & txtMC_Group_Knt.Text & "','" & _Quality & "','" & str20Class & "','" & strSales_Order & "'," & strLine_Item & ",'" & txtQty_Knt.Text & "','" & txtK_Qty.Text & "'," & _WeekNo & "," & Year(_EndDate) & ",'" & _FromDate & "','" & _Todate & "','" & _StartDate & "','" & _EndDate & "','" & txtQuality_Group.Text & "','" & _Qstatus & "')"
                            'up_GetSetBlock_KntPlanning_Boad(ncQryType, nvcFieldList1, vcWhere, connection, transaction)

                            'transaction.Commit()
                            vcWhere = "tmpMC_No='" & M01.Tables(0).Rows(i)("tmpMC_No") & "' and tmpSales_Order='" & strSales_Order & "' and tmpLine_Item='" & strLine_Item & "'"
                            dsUser = DBEngin.ExecuteDataset(connection, transaction, "up_GetSettmp_Knt_Mc", New SqlParameter("@cQryType", "CLS1"), New SqlParameter("@vcWhereClause1", vcWhere))
                            If isValidDataset(dsUser) Then
                                i = i + 1
                                Continue For
                            End If
                            Dim newRow As DataRow = c_dataCustomer1_KNT.NewRow

                            newRow("Machine No") = M01.Tables(0).Rows(i)("tmpMC_No")
                            newRow("Start Date") = _FromDate
                            newRow("End Date") = _Todate
                            Value = txtK_Qty.Text
                            strQty = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                            strQty = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                            newRow("Qty") = strQty
                            newRow("No of Hour") = (_TimeSpam.Days * 24 + _TimeSpam.Hours) & "." & _TimeSpam.Minutes
                            If _Qstatus = "S" Then
                                newRow("Status") = "Same Quality"
                            Else

                                newRow("Status") = "Quality Change"
                            End If
                            newRow("##") = False
                            c_dataCustomer1_KNT.Rows.Add(newRow)

                            Call Update_Tempary_Knt_Mc(M01.Tables(0).Rows(i)("tmpMC_No"), "BADD")
                            'Exit Function
                        Else
                            _Balance_Qty = _Balance_Qty - _TotalTime

                            '' If txtAlocate_MC.Text < _AllocateMC Then
                            'MsgBox("No Available Machine Capacity. Please add the allocate machine", MsgBoxStyle.Information, "Information ....")
                            'vcWhere = "tmpRef_No=" & Delivary_Ref & ""
                            'M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "DKP"), New SqlParameter("@vcWhereClause1", vcWhere))
                            ''transaction.Commit()
                            ''connection.Close()
                            'Exit Function
                            'Else
                            _StartDate = (Month(_FromDate) & "/" & Microsoft.VisualBasic.Day(_FromDate) & "/" & Year(_FromDate))
                            _EndDate = (Month(_Todate) & "/" & Microsoft.VisualBasic.Day(_Todate) & "/" & Year(_Todate))
                            _WeekNo = DatePart("WW", _EndDate, FirstDayOfWeek.Monday)


                            'DEACTIVE BY SURANGA ON 2016.6.17
                            'ncQryType = "KADD"
                            'nvcFieldList1 = "(tmpRef_No," & "tmpMC_No," & "tmpGroup," & "tmpQuality," & "tmp20Class," & "tmpSales_Order," & "tmpLine_Item," & "tmpKnt_Order," & "tmpBalance," & "tmpWeek_No," & "tmpYear," & "tmpStart_Time," & "tmpEnd_Time," & "tmpStart_Date," & "tmpEnd_Date," & "tmpStatus," & "tmpQ_Status) " & "values(" & Delivary_Ref & ",'" & M01.Tables(0).Rows(i)("tmpMC_No") & "','" & txtMC_Group_Knt.Text & "','" & txtQuality.Text & "','" & str20Class & "','" & strSales_Order & "'," & strLine_Item & ",'" & txtQty_Knt.Text & "','" & _TotalTime & "'," & _WeekNo & "," & Year(_EndDate) & ",'" & _FromDate & "','" & _Todate & "','" & _StartDate & "','" & _EndDate & "','" & txtQuality_Group.Text & "','QC')"
                            'up_GetSetBlock_KntPlanning_Boad(ncQryType, nvcFieldList1, vcWhere, connection, transaction)

                            ' transaction.Commit()
                            vcWhere = "tmpMC_No='" & M01.Tables(0).Rows(i)("tmpMC_No") & "' and tmpStart_Time>='" & _FromDate & "' and tmpEnd_Time<='" & _Todate & "'"
                            dsUser = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "KPBT"), New SqlParameter("@vcWhereClause1", vcWhere))
                            If isValidDataset(dsUser) Then
                                i = i + 1
                                Continue For
                            Else

                                vcWhere = "tmpMC_No='" & M01.Tables(0).Rows(i)("tmpMC_No") & "' and tmpSales_Order='" & strSales_Order & "' and tmpLine_Item='" & strLine_Item & "'"
                                dsUser = DBEngin.ExecuteDataset(connection, transaction, "up_GetSettmp_Knt_Mc", New SqlParameter("@cQryType", "CLS1"), New SqlParameter("@vcWhereClause1", vcWhere))
                                If isValidDataset(dsUser) Then
                                    i = i + 1
                                    Continue For
                                End If
                                Dim newRow As DataRow = c_dataCustomer1_KNT.NewRow

                                newRow("Machine No") = M01.Tables(0).Rows(i)("tmpMC_No")
                                newRow("Start Date") = _FromDate
                                newRow("End Date") = _Todate
                                Value = _TotalTime
                                strQty = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                                strQty = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                                newRow("Qty") = strQty
                                newRow("No of Hour") = (_TimeSpam.Days * 24 + _TimeSpam.Hours) & "." & _TimeSpam.Minutes
                                newRow("Status") = "Quality Change"
                                newRow("##") = False
                                c_dataCustomer1_KNT.Rows.Add(newRow)
                                _AllocateMC = _AllocateMC + 1

                                Call Update_Tempary_Knt_Mc(M01.Tables(0).Rows(i)("tmpMC_No"), "BADD")
                                ' End If
                            End If
                        End If
                        i = i + 1
                    Next
                Else

                    Dim userDate As Date
                    Dim WeekEnd_Days As Date
                    _Balance_Qty = _Balance_Qty - _TotalTime

                    '' If txtAlocate_MC.Text < _AllocateMC Then
                    'MsgBox("No Available Machine Capacity. Please add the allocate machine", MsgBoxStyle.Information, "Information ....")
                    'vcWhere = "tmpRef_No=" & Delivary_Ref & ""
                    'M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "DKP"), New SqlParameter("@vcWhereClause1", vcWhere))
                    ''transaction.Commit()
                    ''connection.Close()
                    'Exit Function
                    'Else
                    _FromDate = CDate(txtComplete_Date_Knt.Text).AddDays(-7)
                    If _FromDate > Today Then
                    Else
                        _FromDate = Today
                    End If
                    _StartDate = (Month(_FromDate) & "/" & Microsoft.VisualBasic.Day(_FromDate) & "/" & Year(_FromDate))
                    _StartDate = _StartDate & " " & "7:30AM"

                    userDate = "1/1/" & Year(_StartDate)
                    ' MsgBox(WeekdayName(Weekday(userDate)))
                    If WeekdayName(Weekday(userDate)) = "Sunday" Then
                        userDate = userDate.AddDays(-3)
                    ElseIf WeekdayName(Weekday(userDate)) = "Monday" Then
                        userDate = userDate.AddDays(-4)
                    ElseIf WeekdayName(Weekday(userDate)) = "Tuesday" Then
                        userDate = userDate.AddDays(-5)
                    ElseIf WeekdayName(Weekday(userDate)) = "Thusday" Then
                        'userDate = userDate.AddDays(-1)
                    ElseIf WeekdayName(Weekday(userDate)) = "Friday" Then
                        userDate = userDate.AddDays(-1)
                    ElseIf WeekdayName(Weekday(userDate)) = "Saturday" Then
                        userDate = userDate.AddDays(-2)

                    End If
                    If IsNumeric(txtWeek_Knt.Text) Then
                        _StartDate = userDate.AddDays(txtWeek_Knt.Text * 7)
                    Else
                        Exit Function
                    End If
                    _StartDate = _StartDate & " " & "7:30AM"
                    _StartDate = _StartDate.AddDays(-7)
                    WeekEnd_Days = _StartDate.AddDays(+7)
                    vcWhere = "M40Group_Name='" & Trim(txtQuality_Group.Text) & "' "
                    M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "MGN"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M01) Then
                        _StartDate = _StartDate.AddHours(+M01.Tables(0).Rows(0)("M40Mc_Change_HR"))
                    End If

                    _DAYKNIT = txtK_Qty.Text / txtDaily_Capacity.Text
                    _MinQty = txtK_Qty.Text / txtDaily_Capacity.Text
                    _MinQty = _MinQty * 24 * 60
                    ' MsgBox(CInt(_MinQty))
                    _MinQty = _MinQty + 1
                    _Todate = _StartDate.AddMinutes(_MinQty)
                    If _Todate > WeekEnd_Days Then
                        _Todate = WeekEnd_Days
                    End If
                    _TimeSpam = _Todate.Subtract(_StartDate)
                    _TotalTime = (_TimeSpam.Days * 24 * 60)
                    _TotalTime = _TotalTime + (_TimeSpam.Hours * 60)
                    _TotalTime = _TotalTime + _TimeSpam.Minutes
                    _TotalTime = _TotalTime * _Qty_Minite
                    _EndDate = (Month(_Todate) & "/" & Microsoft.VisualBasic.Day(_Todate) & "/" & Year(_Todate))
                    _WeekNo = DatePart("WW", _EndDate, FirstDayOfWeek.Monday)

                    vcWhere = "M38MC='" & _McName & "' and M38Group='" & strGuarge & "'"
                    M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "EQA"), New SqlParameter("@vcWhereClause1", vcWhere))
                    i = 0
                    For Each DTRow4 As DataRow In M01.Tables(0).Rows
                        'If _Balance_Qty <= 0 Then
                        '    transaction.Commit()
                        Call Calculate_Tobe_Knt_Balance()
                        '    Exit Function
                        'End If
                        'If _Balance_Qty > _TotalTime Then
                        '    Value = _TotalTime
                        '    'newRow("Start Date") = _StartDate
                        '    'newRow("End Date") = _Todate
                        'Else

                        '    Value = _Balance_Qty
                        ' End If
                        vcWhere = "tmpMC_No='" & M01.Tables(0).Rows(i)("M39MC_NO") & "' and tmpEnd_Time>='" & _StartDate & "' "
                        dsUser = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "SRK"), New SqlParameter("@vcWhereClause1", vcWhere))
                        If isValidDataset(dsUser) Then
                            i = i + 1
                            Continue For
                        End If

                        vcWhere = "tmpMC_No='" & M01.Tables(0).Rows(i)("M39MC_NO") & "' and tmpSales_Order='" & strSales_Order & "' and tmpLine_Item='" & strLine_Item & "'"
                        dsUser = DBEngin.ExecuteDataset(connection, transaction, "up_GetSettmp_Knt_Mc", New SqlParameter("@cQryType", "CLS1"), New SqlParameter("@vcWhereClause1", vcWhere))
                        If isValidDataset(dsUser) Then
                            i = i + 1
                            Continue For
                        End If

                        'DEACTIVE BY SURANGA ON 2016.6.17
                        'ncQryType = "KADD"
                        'nvcFieldList1 = "(tmpRef_No," & "tmpMC_No," & "tmpGroup," & "tmpQuality," & "tmp20Class," & "tmpSales_Order," & "tmpLine_Item," & "tmpKnt_Order," & "tmpBalance," & "tmpWeek_No," & "tmpYear," & "tmpStart_Time," & "tmpEnd_Time," & "tmpStart_Date," & "tmpEnd_Date," & "tmpStatus," & "tmpQ_Status) " & "values(" & Delivary_Ref & ",'" & M01.Tables(0).Rows(i)("M39MC_NO") & "','" & txtMC_Group_Knt.Text & "','" & txtQuality.Text & "','" & str20Class & "','" & strSales_Order & "'," & strLine_Item & ",'" & Value & "','" & CDbl(lblKnt_Balance.Text) - txtK_Qty.Text & "'," & _WeekNo & "," & Year(_EndDate) & ",'" & _FromDate & "','" & _Todate & "','" & _StartDate & "','" & _EndDate & "','" & txtQuality_Group.Text & "','QC')"
                        'up_GetSetBlock_KntPlanning_Boad(ncQryType, nvcFieldList1, vcWhere, connection, transaction)


                        ' transaction.Commit()

                        Dim newRow As DataRow = c_dataCustomer1_KNT.NewRow

                        newRow("Machine No") = M01.Tables(0).Rows(i)("M39MC_NO")

                        ' If _Balance_Qty > _TotalTime Then
                        Value = _TotalTime
                        newRow("Start Date") = _StartDate
                        newRow("End Date") = _Todate
                        'Else

                        'Value = _Balance_Qty
                        'newRow("Start Date") = _StartDate
                        ''_TotalTime = _Balance_Qty / txtDaily_Capacity.Text
                        ''_Todate = _StartDate.AddDays(+Fix(_TotalTime))
                        ''_TotalTime = _Balance_Qty - (Fix(_TotalTime) * txtDaily_Capacity.Text)
                        ''_TotalTime = _TotalTime * _MinQty
                        ''_Todate = _Todate.AddMinutes(+Fix(_TotalTime))
                        ''_TotalTime = _Balance_Qty
                        'newRow("End Date") = _Todate
                        'End If

                        _TimeSpam = _Todate.Subtract(_StartDate)
                        strQty = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        strQty = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                        newRow("Qty") = strQty
                        newRow("No of Hour") = (_TimeSpam.Days * 24 + _TimeSpam.Hours) & "." & _TimeSpam.Minutes
                        newRow("Status") = "Quality Change"
                        newRow("##") = False
                        c_dataCustomer1_KNT.Rows.Add(newRow)

                        Update_Tempary_Knt_Mc(M01.Tables(0).Rows(i)("M39MC_NO"), "BADD")

                        '  _Balance_Qty = _Balance_Qty - _TotalTime
                        'If _Balance_Qty > 0 Then
                        'Else
                        '    _Balance_Qty = 0
                        'End If
                        i = i + 1
                    Next
                End If
                x = x + 1
            Next
            'End If
            ' transaction.Commit()
            connection.Close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.Close()
            End If
        End Try

    End Function

    Private Sub txtWeek_Knt_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtWeek_Knt.KeyUp
        If e.KeyCode = 13 Then
            txtYear_Knt.Focus()
        ElseIf e.KeyCode = Keys.Tab Then
            txtYear_Knt.Focus()
        ElseIf e.KeyCode = Keys.F1 Then
            OPR_Kplan.Visible = True
        End If
    End Sub

    Private Sub txtYear_Knt_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtYear_Knt.KeyUp
        If e.KeyCode = 13 Then
            txtComplete_Date_Knt.Focus()
            Call TestDateAdd()
        ElseIf e.KeyCode = Keys.Tab Then
            txtComplete_Date_Knt.Focus()
            Call TestDateAdd()
        ElseIf e.KeyCode = Keys.F1 Then
            OPR_Kplan.Visible = True
        End If
    End Sub

    Private Sub txtYear_Knt_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtYear_Knt.LostFocus
        ' Call TestDateAdd()
    End Sub

    Private Sub TestDateAdd()
        If IsNumeric(txtWeek_Knt.Text) And IsNumeric(txtYear_Knt.Text) Then
            Dim weekStart As DateTime = GetWeekStartDate(txtWeek_Knt.Text, txtYear_Knt.Text)
            Dim _Wk As Integer

            weekStart = "1/1/" & txtYear_Knt.Text
            If WeekdayName(Weekday(weekStart)) = "Sunday" Then
                weekStart = weekStart.AddDays(-3)
            ElseIf WeekdayName(Weekday(weekStart)) = "Monday" Then
                weekStart = weekStart.AddDays(-4)
            ElseIf WeekdayName(Weekday(weekStart)) = "Tuesday" Then
                weekStart = weekStart.AddDays(-5)
            ElseIf WeekdayName(Weekday(weekStart)) = "Thusday" Then
                'userDate = userDate.AddDays(-1)
            ElseIf WeekdayName(Weekday(weekStart)) = "Friday" Then
                weekStart = weekStart.AddDays(-1)
            ElseIf WeekdayName(Weekday(weekStart)) = "Saturday" Then
                weekStart = weekStart.AddDays(-2)

            End If

            _Wk = 7 * txtWeek_Knt.Text
            ' _Wk = _Wk - 1
            weekStart = weekStart.AddDays(+_Wk)
            ' weekStart = weekStart.AddDays(-7)
            txtComplete_Date_Knt.Text = weekStart
        End If
    End Sub

    Private Function GetWeekStartDate(ByVal weekNumber As Integer, ByVal year As Integer) As Date
        Dim startDate As New DateTime(year, 1, 1)
        Dim weekDate As DateTime = DateAdd(DateInterval.WeekOfYear, weekNumber - 1, startDate)
        Return DateAdd(DateInterval.Day, (-weekDate.DayOfWeek) + 1, weekDate)
    End Function

    Private Sub txtComplete_Date_Knt_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtComplete_Date_Knt.TextChanged
        Call Search_WeekNo()
    End Sub

    Private Sub txtUse_FG_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtUse_FG.KeyUp
        Dim Value As Double

        If e.KeyCode = 13 Then
            Call frmDelivaryQuatnew.CalculateBalance_To_Produce()
            If IsNumeric(txtUse_FG.Text) Then
                Value = txtUse_FG.Text
                txtUse_FG.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtUse_FG.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            End If
            txtUse_WIP.Focus()
        ElseIf e.KeyCode = Keys.Tab Then
            Call frmDelivaryQuatnew.CalculateBalance_To_Produce()
            If IsNumeric(txtUse_FG.Text) Then
                Value = txtUse_FG.Text
                txtUse_FG.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtUse_FG.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            End If
            txtUse_WIP.Focus()

        End If
    End Sub

    Private Sub txtUse_FG_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtUse_FG.TextChanged
        On Error Resume Next
        Dim _Balance As Double
        Call frmDelivaryQuatnew.CalculateBalance_To_Produce()

        If IsNumeric(txtConfact.Text) And IsNumeric(txtDye_Wast.Text) Then
            _Balance = CDbl(txtBalance.Text) / CDbl(txtConfact.Text)
            _Balance = _Balance / ((100 - CDbl(txtDye_Wast.Text)) / 100)
            '_Balance = _Balance / 100
            _Balance = _Balance '+ CDbl(txtBalance.Text)
            txtReq_Grg.Text = (_Balance.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            txtReq_Grg.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", _Balance))
        Else
            _Balance = CDbl(txtBalance.Text)
            txtReq_Grg.Text = (_Balance.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            txtReq_Grg.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", _Balance))
        End If

    End Sub

    Private Sub txtUse_WIP_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtUse_WIP.KeyUp
        Dim Value As Double
        If e.KeyCode = 13 Then
            Call frmDelivaryQuatnew.CalculateBalance_To_Produce()
            If IsNumeric(txtUse_WIP.Text) Then
                Value = txtUse_WIP.Text
                txtUse_WIP.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtUse_WIP.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            End If

        ElseIf e.KeyCode = Keys.Tab Then
            Call frmDelivaryQuatnew.CalculateBalance_To_Produce()
            If IsNumeric(txtUse_WIP.Text) Then
                Value = txtUse_WIP.Text
                txtUse_WIP.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtUse_WIP.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            End If
        End If
    End Sub
    Private Sub txtUse_WIP_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtUse_WIP.TextChanged
        Call frmDelivaryQuatnew.CalculateBalance_To_Produce()
    End Sub

    Private Sub txtEx_LibUse_LostFocus1(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEx_LibUse.LostFocus
        Call frmDelivaryQuatnew.CalculateBalance_To_Produce()
    End Sub

   
    Private Sub UltraButton9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton9.Click
        Call Load_Gride_YB()
        Call Load_Gridewith_DataSerch_YB()
    End Sub

   
    Private Sub chkCh1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkCh1.CheckedChanged
        If chkCh1.Checked = True Then
            chkCh2.Checked = False
            chkCh3.Checked = False
        End If
    End Sub

    Private Sub chkCh2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkCh2.CheckedChanged
        If chkCh2.Checked = True Then
            chkCh1.Checked = False
            chkCh3.Checked = False
        End If
    End Sub

    Private Sub chkCh3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkCh3.CheckedChanged
        If chkCh3.Checked = True Then
            chkCh2.Checked = False
            chkCh1.Checked = False
        End If
    End Sub

   

    Function Update_Records_YARN_REQUEST()
        Dim vcWhere As String
        Dim M01 As DataSet
        Dim vcFieldList As String
        Dim ncQryType As String
        Dim nvcFieldList1 As String
        Dim M02 As DataSet

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Dim i As Integer
        Dim _Balance As Double
        i = 0
        '1st Yarn
        Try
            If txtYarn1.Text <> "" Then
                For Each uRow As UltraGridRow In dg1_YB.Rows
                    If Microsoft.VisualBasic.Left(Trim(dg1_YB.Rows(i).Cells(1).Text), 4) = Microsoft.VisualBasic.Left(txtYarn1.Text, 4) Then
                        If Trim(dg1_YB.Rows(i).Cells(6).Text) <> "" Then
                            _Balance = _Balance + Trim(dg1_YB.Rows(i).Cells(6).Value)
                        End If
                    End If

                    i = i + 1
                Next
            End If

            If (CDbl(txtReq1.Text) - _Balance) > 0 Then
                vcWhere = "T14Ref_no=" & Delivary_Ref & " and T14Sales_order='" & strSales_Order & "' and T14Line_Item=" & strLine_Item & " and T14Yarn='" & txtYarn1.Text & "'"
                M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "YRQ"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then

                Else
                    ncQryType = "AYR"
                    nvcFieldList1 = "(T14Ref_no," & "T14Sales_order," & "T14Line_Item," & "T14Yarn," & "T14Req_By," & "T14Req_Date," & "T14Status," & "T14Qty," & "T14Time) " & "values(" & Delivary_Ref & ",'" & strSales_Order & "'," & strLine_Item & ",'" & txtYarn1.Text & "','" & strDisname & "','" & Today & "','N','" & CDbl(txtReq1.Text) - _Balance & "','" & Now & "')"
                    up_GetSetYarn_Request(ncQryType, nvcFieldList1, vcWhere, connection, transaction)


                End If
            End If
            '2nd Yarn
            i = 0
            _Balance = 0
            If txtYarn2.Text <> "" Then
                For Each uRow As UltraGridRow In dg1_YB.Rows
                    If Microsoft.VisualBasic.Left(Trim(dg1_YB.Rows(i).Cells(1).Text), 4) = Microsoft.VisualBasic.Left(txtYarn2.Text, 4) Then
                        If Trim(dg1_YB.Rows(i).Cells(6).Text) <> "" Then
                            _Balance = _Balance + Trim(dg1_YB.Rows(i).Cells(6).Value)
                        End If
                    End If

                    i = i + 1
                Next
            End If

            If (CDbl(txtReq2.Text) - _Balance) > 0 Then
                vcWhere = "T14Ref_no=" & Delivary_Ref & " and T14Sales_order='" & strSales_Order & "' and T14Line_Item=" & strLine_Item & " and T14Yarn='" & txtYarn2.Text & "'"
                M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "YRQ"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then

                Else
                    ncQryType = "AYR"
                    nvcFieldList1 = "(T14Ref_no," & "T14Sales_order," & "T14Line_Item," & "T14Yarn," & "T14Req_By," & "T14Req_Date," & "T14Status," & "T14Qty," & "T14Time) " & "values(" & Delivary_Ref & ",'" & strSales_Order & "'," & strLine_Item & ",'" & txtYarn2.Text & "','" & strDisname & "','" & Today & "','N','" & CDbl(txtReq2.Text) - _Balance & "','" & Now & "')"
                    up_GetSetYarn_Request(ncQryType, nvcFieldList1, vcWhere, connection, transaction)


                End If
            End If

            transaction.Commit()
            connection.Close()
            Me.Close()

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)

                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
                connection.Close()
            End If
        End Try
    End Function

    Function Update_Records_DYEDYARN_REQUEST()
        Dim vcWhere As String
        Dim M01 As DataSet
        Dim vcFieldList As String
        Dim ncQryType As String
        Dim nvcFieldList1 As String
        Dim M02 As DataSet

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Dim i As Integer
        Dim Z As Integer
        Dim _Balance As Double
        i = 0
        '1st Yarn
        Try

         
            Z = 0
            For Each uRow1 As UltraGridRow In dgDYYarn_Request.Rows
                i = 0
                _Balance = 0
                For Each uRow As UltraGridRow In dg1.Rows
                    If Trim(dg1.Rows(i).Cells(0).Text) <> "" Then
                        If Trim(dg1.Rows(i).Cells(1).Text) = Trim(dgDYYarn_Request.Rows(Z).Cells(0).Text) Then
                            If Trim(dg1.Rows(i).Cells(9).Text) <> "" Then
                                _Balance = Trim(dg1.Rows(i).Cells(7).Value) - Trim(dg1.Rows(i).Cells(9).Value)
                            Else
                                _Balance = Trim(dg1.Rows(i).Cells(7).Value)
                            End If
                            vcWhere = "T14Ref_no=" & Delivary_Ref & " and T14Sales_order='" & strSales_Order & "' and T14Line_Item=" & strLine_Item & " and T14Yarn='" & Trim(dg1.Rows(i).Cells(1).Text) & "'"
                            M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "YRQ"), New SqlParameter("@vcWhereClause1", vcWhere))
                            If isValidDataset(M01) Then
                                nvcFieldList1 = "UPDATE T14Yarn_Request SET T14Capacity='" & CDbl(dgDYYarn_Request.Rows(Z).Cells(3).Text) & "',T14Available='" & _Balance & "' WHERE T14Ref_no=" & Delivary_Ref & " and T14Sales_order='" & strSales_Order & "' and T14Line_Item=" & strLine_Item & " and T14Yarn='" & Trim(dg1.Rows(i).Cells(1).Text) & "'"
                                up_GetSetYarn_Request(ncQryType, nvcFieldList1, vcWhere, connection, transaction)

                            Else
                                ncQryType = "AYR"
                                nvcFieldList1 = "(T14Ref_no," & "T14Sales_order," & "T14Line_Item," & "T14Yarn," & "T14Req_By," & "T14Year," & "T14Status," & "T14Capacity," & "T14Time," & "T14Class," & "T14Week," & "T14Available) " & "values(" & Delivary_Ref & ",'" & strSales_Order & "'," & strLine_Item & ",'" & Trim(dg1.Rows(i).Cells(1).Text) & "','" & strDisname & "','" & Trim(dgDYYarn_Request.Rows(Z).Cells(1).Text) & "','N','" & CDbl(dgDYYarn_Request.Rows(Z).Cells(3).Text) & "','" & Now & "','" & Trim(dg1.Rows(i).Cells(0).Text) & "','" & Trim(dgDYYarn_Request.Rows(Z).Cells(2).Text) & "','" & _Balance & "')"
                                up_GetSetYarn_Request(ncQryType, nvcFieldList1, vcWhere, connection, transaction)

                            End If
                        End If
                    End If
                    i = i + 1
                Next

                Z = Z + 1
            Next
            
            '=============================================================
            'UPDATE PLANE DETAILES & GRIGE STOCK
            nvcFieldList1 = "DELETE FROM tmpYarn_Booking WHERE tmpRef=" & Delivary_Ref & ""
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)



            ' MsgBox("Send Yarn Request Successfully", MsgBoxStyle.Information, "Information ......")
            transaction.Commit()
            connection.Close()
            'Me.Close()

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)

                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
                connection.Close()
            End If
        End Try
    End Function

    Function Update_Records_YARN_REQUEST_PROCUMENT()
        Dim vcWhere As String
        Dim M01 As DataSet
        Dim vcFieldList As String
        Dim ncQryType As String
        Dim nvcFieldList1 As String
        Dim M02 As DataSet

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean
        Dim _10Class As String

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Dim i As Integer
        Dim Z As Integer
        Dim _Balance As Double
        i = 0
        '1st Yarn
        Try


            Z = 0
            For Each uRow1 As UltraGridRow In dg_Yarn_Request.Rows
                'YARN 1
                If Trim(lblBalance_YB.Text) <> "" Then
                    If Trim(txtYarn1.Text) = Trim(dg_Yarn_Request.Rows(Z).Cells(0).Text) Then
                        _Balance = lblBalance_YB.Text
                    End If

                    _10Class = "-"


                    vcWhere = "T14Ref_no=" & Delivary_Ref & " and T14Sales_order='" & strSales_Order & "' and T14Line_Item=" & strLine_Item & " and T14Yarn='" & txtYarn1.Text & "'"
                    M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "YRQ"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M01) Then
                        nvcFieldList1 = "UPDATE T14Yarn_Request SET T14Capacity='" & Trim(dg_Yarn_Request.Rows(Z).Cells(3).Text) & "',T14Available='" & _Balance & "' WHERE T14Ref_no=" & Delivary_Ref & " and T14Sales_order='" & strSales_Order & "' and T14Line_Item=" & strLine_Item & " and T14Yarn='" & txtYarn1.Text & "'"
                        up_GetSetYarn_Request(ncQryType, nvcFieldList1, vcWhere, connection, transaction)

                    Else
                        ncQryType = "AYR"
                        nvcFieldList1 = "(T14Ref_no," & "T14Sales_order," & "T14Line_Item," & "T14Yarn," & "T14Req_By," & "T14Year," & "T14Status," & "T14Capacity," & "T14Time," & "T14Class," & "T14Week," & "T14Available) " & "values(" & Delivary_Ref & ",'" & strSales_Order & "'," & strLine_Item & ",'" & txtYarn1.Text & "','" & strDisname & "','" & Trim(dg_Yarn_Request.Rows(Z).Cells(1).Text) & "','N','" & CDbl(dg_Yarn_Request.Rows(Z).Cells(3).Text) & "','" & Now & "','" & _10Class & "','" & Trim(dg_Yarn_Request.Rows(Z).Cells(2).Text) & "','" & _Balance & "')"
                        up_GetSetYarn_Request(ncQryType, nvcFieldList1, vcWhere, connection, transaction)

                    End If
                    Z = Z + 1
                    Continue For
                End If
                'YARN 2
                If Trim(lblBalance_YB1.Text) <> "" Then
                    If Trim(txtYarn2.Text) = Trim(dg_Yarn_Request.Rows(Z).Cells(0).Text) Then
                        _Balance = lblBalance_YB1.Text
                    End If

                    _10Class = "-"


                    vcWhere = "T14Ref_no=" & Delivary_Ref & " and T14Sales_order='" & strSales_Order & "' and T14Line_Item=" & strLine_Item & " and T14Yarn='" & txtYarn2.Text & "'"
                    M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "YRQ"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M01) Then
                        nvcFieldList1 = "UPDATE T14Yarn_Request SET T14Capacity='" & Trim(dg_Yarn_Request.Rows(Z).Cells(3).Text) & "',T14Available='" & _Balance & "' WHERE T14Ref_no=" & Delivary_Ref & " and T14Sales_order='" & strSales_Order & "' and T14Line_Item=" & strLine_Item & " and T14Yarn='" & txtYarn2.Text & "'"
                        up_GetSetYarn_Request(ncQryType, nvcFieldList1, vcWhere, connection, transaction)

                    Else
                        ncQryType = "AYR"
                        nvcFieldList1 = "(T14Ref_no," & "T14Sales_order," & "T14Line_Item," & "T14Yarn," & "T14Req_By," & "T14Year," & "T14Status," & "T14Capacity," & "T14Time," & "T14Class," & "T14Week," & "T14Available) " & "values(" & Delivary_Ref & ",'" & strSales_Order & "'," & strLine_Item & ",'" & txtYarn2.Text & "','" & strDisname & "','" & Trim(dg_Yarn_Request.Rows(Z).Cells(1).Text) & "','N','" & CDbl(dg_Yarn_Request.Rows(Z).Cells(3).Text) & "','" & Now & "','" & _10Class & "','" & Trim(dg_Yarn_Request.Rows(Z).Cells(2).Text) & "','" & _Balance & "')"
                        up_GetSetYarn_Request(ncQryType, nvcFieldList1, vcWhere, connection, transaction)

                    End If
                    Z = Z + 1
                    Continue For
                End If

                'YARN 3
                If Trim(lblBalance_YB2.Text) <> "" Then
                    If Trim(txtYarn2.Text) = Trim(dg_Yarn_Request.Rows(Z).Cells(0).Text) Then
                        _Balance = lblBalance_YB2.Text
                    End If

                    _10Class = "-"


                    vcWhere = "T14Ref_no=" & Delivary_Ref & " and T14Sales_order='" & strSales_Order & "' and T14Line_Item=" & strLine_Item & " and T14Yarn='" & txtYarn2.Text & "'"
                    M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "YRQ"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M01) Then
                        nvcFieldList1 = "UPDATE T14Yarn_Request SET T14Capacity='" & Trim(dg_Yarn_Request.Rows(Z).Cells(3).Text) & "',T14Available='" & _Balance & "' WHERE T14Ref_no=" & Delivary_Ref & " and T14Sales_order='" & strSales_Order & "' and T14Line_Item=" & strLine_Item & " and T14Yarn='" & txtYarn2.Text & "'"
                        up_GetSetYarn_Request(ncQryType, nvcFieldList1, vcWhere, connection, transaction)

                    Else
                        ncQryType = "AYR"
                        nvcFieldList1 = "(T14Ref_no," & "T14Sales_order," & "T14Line_Item," & "T14Yarn," & "T14Req_By," & "T14Year," & "T14Status," & "T14Capacity," & "T14Time," & "T14Class," & "T14Week," & "T14Available) " & "values(" & Delivary_Ref & ",'" & strSales_Order & "'," & strLine_Item & ",'" & txtYarn2.Text & "','" & strDisname & "','" & Trim(dg_Yarn_Request.Rows(Z).Cells(1).Text) & "','N','" & CDbl(dg_Yarn_Request.Rows(Z).Cells(3).Text) & "','" & Now & "','" & _10Class & "','" & Trim(dg_Yarn_Request.Rows(Z).Cells(2).Text) & "','" & _Balance & "')"
                        up_GetSetYarn_Request(ncQryType, nvcFieldList1, vcWhere, connection, transaction)

                    End If
                    Z = Z + 1
                    Continue For
                End If
                Z = Z + 1
            Next

            '=============================================================
            'UPDATE PLANE DETAILES & GRIGE STOCK
            nvcFieldList1 = "DELETE FROM tmpYarn_Booking WHERE tmpRef=" & Delivary_Ref & ""
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)



            ' MsgBox("Send Yarn Request Successfully", MsgBoxStyle.Information, "Information ......")
            transaction.Commit()
            connection.Close()
            'Me.Close()

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)

                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
                connection.Close()
            End If
        End Try
    End Function

    Private Sub cmdEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Call Update_Records_YARN_REQUEST()
    End Sub


    Function MakeDataTable_Greige_Rady() As DataTable
        Dim i As Integer
        Dim vcWhere As String
        Dim M01 As DataSet
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim strWeek As String

        Try
            ' declare a DataTable to contain the program generated data
            Dim dataTable As New DataTable("StkItem")
            ' create and add a Code column
            Dim colWork As New DataColumn("##", GetType(Boolean))
            dataTable.Columns.Add(colWork)
            '' add CustomerID column to key array and bind to DataTable
            ' Dim Keys(0) As DataColumn

            colWork = New DataColumn("Line Item", GetType(String))
            colWork.MaxLength = 250
            dataTable.Columns.Add(colWork)
            colWork.ReadOnly = True

            colWork = New DataColumn("Material", GetType(String))
            colWork.MaxLength = 250
            dataTable.Columns.Add(colWork)
            colWork.ReadOnly = True


            colWork = New DataColumn("Description", GetType(String))
            colWork.MaxLength = 250
            dataTable.Columns.Add(colWork)
            colWork.ReadOnly = True

            colWork = New DataColumn("Quantity", GetType(String))
            colWork.MaxLength = 250
            dataTable.Columns.Add(colWork)


            colWork = New DataColumn("In Stock", GetType(String))
            colWork.MaxLength = 120
            dataTable.Columns.Add(colWork)
            colWork.ReadOnly = True

            i = 0
            vcWhere = "tmpSales_Order='" & strSales_Order & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "SRK"), New SqlParameter("@vcWhereClause1", vcWhere))
            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                strWeek = "Week " & M01.Tables(0).Rows(i)("tmpWeek_No")
                colWork = New DataColumn(strWeek, GetType(String))
                colWork.MaxLength = 120
                dataTable.Columns.Add(colWork)
                colWork.ReadOnly = True
                i = i + 1
            Next


            Return dataTable


            con.Close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.Close()
            End If
        End Try
    End Function

    Function MakeDataTable_Greige_Rady_Details() As DataTable
        '2nd Screen
        Dim i As Integer
        Dim vcWhere As String
        Dim M01 As DataSet
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim strWeek As String

        Try
            ' declare a DataTable to contain the program generated data
            Dim dataTable As New DataTable("StkItem")
            ' create and add a Code column
            Dim colWork As New DataColumn("##", GetType(String))
            dataTable.Columns.Add(colWork)
            '' add CustomerID column to key array and bind to DataTable
            ' Dim Keys(0) As DataColumn

            colWork = New DataColumn("Line Item", GetType(String))
            colWork.MaxLength = 250
            dataTable.Columns.Add(colWork)
            colWork.ReadOnly = True

            colWork = New DataColumn("Material", GetType(String))
            colWork.MaxLength = 250
            dataTable.Columns.Add(colWork)
            colWork.ReadOnly = True


            colWork = New DataColumn("Description", GetType(String))
            colWork.MaxLength = 250
            dataTable.Columns.Add(colWork)
            colWork.ReadOnly = True

            colWork = New DataColumn("Quantity", GetType(String))
            colWork.MaxLength = 250
            dataTable.Columns.Add(colWork)


            colWork = New DataColumn("In Stock", GetType(String))
            colWork.MaxLength = 120
            dataTable.Columns.Add(colWork)
            colWork.ReadOnly = True

            i = 0
            vcWhere = "tmpSales_Order='" & strSales_Order & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "SRK"), New SqlParameter("@vcWhereClause1", vcWhere))
            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                strWeek = "Week " & M01.Tables(0).Rows(i)("tmpWeek_No")
                colWork = New DataColumn(strWeek, GetType(String))
                colWork.MaxLength = 120
                dataTable.Columns.Add(colWork)
                colWork.ReadOnly = True
                i = i + 1
            Next


            Return dataTable


            con.Close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.Close()
            End If
        End Try
    End Function

    Function Load_Gride_Dye_Main()
        Dim vcwhere As String
        Dim i As Integer
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim M02 As DataSet
        Dim M03 As DataSet

        Dim Value As Double
        Dim _ST As String
        Dim Y As Integer
        Dim strWeek As String

        Try
            i = 0
            Value = 0
            vcwhere = "tmpSales_Order='" & strSales_Order & "' and tmpQ_Status IS NULL"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "SRK1"), New SqlParameter("@vcWhereClause1", vcwhere))
            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1_Dye.NewRow

                newRow("##") = False
                newRow("Line Item") = M01.Tables(0).Rows(i)("tmpLine_Item")
                newRow("Material") = M01.Tables(0).Rows(i)("M01Material_No")
                newRow("Description") = M01.Tables(0).Rows(i)("M01Quality")
                Value = M01.Tables(0).Rows(i)("Qty")
                _ST = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _ST = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Quantity") = _ST


                Y = 0
                vcwhere = "tmpSales_Order='" & strSales_Order & "' and tmpLine_Item=" & M01.Tables(0).Rows(i)("tmpLine_Item") & " "
                M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "SRK"), New SqlParameter("@vcWhereClause1", vcwhere))
                For Each DTRow5 As DataRow In M02.Tables(0).Rows

                    vcwhere = "tmpSales_Order='" & strSales_Order & "' and tmpLine_Item=" & M01.Tables(0).Rows(i)("tmpLine_Item") & " and tmpWeek_No=" & M02.Tables(0).Rows(Y)("tmpWeek_No") & ""
                    M03 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "SRK1"), New SqlParameter("@vcWhereClause1", vcwhere))
                    If isValidDataset(M03) Then
                        strWeek = "Week " & M02.Tables(0).Rows(Y)("tmpWeek_No")
                        Value = M03.Tables(0).Rows(0)("Qty")
                        _ST = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        _ST = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                        newRow(strWeek) = _ST

                    End If
                    Y = Y + 1
                Next

                c_dataCustomer1_Dye.Rows.Add(newRow)
                i = i + 1
            Next

            con.Close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.Close()
            End If
        End Try
    End Function

    Function MakeDataTable_Projection_Dye(ByVal _Quality_No As String) As DataTable

        Dim vcwhere As String
        Dim i As Integer
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim M02 As DataSet
        Dim M03 As DataSet

        Dim Value As Double
        Dim _ST As String
        Dim Y As Integer
        Dim strWeek As String
        Dim _Date As Date
        Dim _Code As Integer
        Dim Z As Integer
       

        Try
            ' declare a DataTable to contain the program generated data
            Dim dataTable As New DataTable("StkItem")
            ' create and add a Code column
            Dim colWork As New DataColumn("Code", GetType(String))
            dataTable.Columns.Add(colWork)
            '' add CustomerID column to key array and bind to DataTable
            ' Dim Keys(0) As DataColumn
            vcwhere = "select * from P01PARAMETER where P01CODE='PRN'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, vcwhere)
            If isValidDataset(M01) Then
                _Code = M01.Tables(0).Rows(0)("P01NO")
            End If

            _Code = _Code - 1
            _Projection_Code = _Code
            If Microsoft.VisualBasic.Day(Today) > 10 Then
                _Date = Today.AddDays(+30)
                _Date = Month(_Date) & "/1/" & Year(_Date)
                vcwhere = "M43Quality in ('" & _Quality_No & "') and convert(datetime,Ddate,111)>='" & _Date & "'  and M43Count_No=" & _Code & ""
            Else
                _Date = Today
                _Date = Month(_Date) & "/1/" & Year(_Date)
                vcwhere = "M43Quality in ('" & _Quality_No & "') and convert(datetime,Ddate,111)>='" & _Date & "'  and M43Count_No=" & _Code & ""
            End If
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "PRS2"), New SqlParameter("@vcWhereClause1", vcwhere))
            i = 0
            For Each DTRow5 As DataRow In M01.Tables(0).Rows
                If i = 0 Then
                    _DyeMonth = M01.Tables(0).Rows(i)("M43Product_Month")
                    _DyeYear = M01.Tables(0).Rows(i)("M43Year")
                Else
                    _DyeMonth = _DyeMonth & "','" & M01.Tables(0).Rows(i)("M43Product_Month")
                    _DyeYear = _DyeYear & "','" & M01.Tables(0).Rows(i)("M43Year")
                End If
                strWeek = M01.Tables(0).Rows(i)("M43Product_Month")
                colWork = New DataColumn(MonthName(strWeek), GetType(String))
                colWork.MaxLength = 250
                dataTable.Columns.Add(colWork)
                ' colWork.ReadOnly = True
                i = i + 1
            Next

            '==================================================================
           

            Return dataTable

            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            con.close()

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)

                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
                con.close()
            End If
        End Try
    End Function

    Function MakeDataTable_Capacity_Dye(ByVal _Quality_No As String) As DataTable

        Dim vcwhere As String
        Dim i As Integer
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim M02 As DataSet
        Dim M03 As DataSet

        Dim Value As Double
        Dim _ST As String
        Dim Y As Integer
        Dim strWeek As String
        Dim _Date As Date
        Dim _Code As Integer
        Dim Z As Integer
        Dim _WeekNo As Integer
        Dim _weekSt As Date


        Try
            ' declare a DataTable to contain the program generated data
            Dim dataTable As New DataTable("StkItem")
            ' create and add a Code column
            Dim colWork As New DataColumn("##", GetType(String))
            dataTable.Columns.Add(colWork)
            '' add CustomerID column to key array and bind to DataTable
            ' Dim Keys(0) As DataColumn
            vcwhere = "select * from P01PARAMETER where P01CODE='PRN'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, vcwhere)
            If isValidDataset(M01) Then
                _Code = M01.Tables(0).Rows(0)("P01NO")
            End If

            _Code = _Code - 1


          

          
            ' colWork.ReadOnly = True

            i = 0
            vcwhere = "T15Sales_Order='" & strSales_Order & "' and T15Line_Item in ('" & _Quality_No & "') "
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "PROX"), New SqlParameter("@vcWhereClause1", vcwhere))
            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                Dim remain As Integer
                Dim noOfWeek As Integer
                Dim userDate As Date
                Dim _LastDate As Date
                Dim _TimeSpan As TimeSpan

                userDate = DateTime.Parse(M01.Tables(0).Rows(i)("T15Month") & "/1/" & M01.Tables(0).Rows(i)("T15Year"))
                ' MsgBox(WeekdayName(Weekday(userDate)))
                If WeekdayName(Weekday(userDate)) = "Sunday" Then
                    userDate = userDate.AddDays(-3)
                ElseIf WeekdayName(Weekday(userDate)) = "Monday" Then
                    userDate = userDate.AddDays(-4)
                ElseIf WeekdayName(Weekday(userDate)) = "Tuesday" Then
                    userDate = userDate.AddDays(-5)
                ElseIf WeekdayName(Weekday(userDate)) = "Thusday" Then
                    'userDate = userDate.AddDays(-1)
                ElseIf WeekdayName(Weekday(userDate)) = "Friday" Then
                    userDate = userDate.AddDays(-1)
                ElseIf WeekdayName(Weekday(userDate)) = "Saturday" Then
                    userDate = userDate.AddDays(-2)

                End If

                _LastDate = DateTime.Parse(M01.Tables(0).Rows(i)("T15Month") & "/1/" & M01.Tables(0).Rows(i)("T15Year"))
                ' MsgBox(Date.DaysInMonth(_LastDate.Year, _LastDate.Month))
                _LastDate = DateTime.Parse(M01.Tables(0).Rows(i)("T15Month") & "/" & Date.DaysInMonth(_LastDate.Year, _LastDate.Month) & "/" & M01.Tables(0).Rows(i)("T15Year"))

                If WeekdayName(Weekday(_LastDate)) = "Sunday" Then
                    _LastDate = _LastDate.AddDays(-4)
                ElseIf WeekdayName(Weekday(_LastDate)) = "Monday" Then
                    _LastDate = _LastDate.AddDays(-5)
                ElseIf WeekdayName(Weekday(_LastDate)) = "Tuesday" Then
                    _LastDate = _LastDate.AddDays(-6)
                ElseIf WeekdayName(Weekday(_LastDate)) = "Thusday" Then
                    'userDate = userDate.AddDays(-1)
                ElseIf WeekdayName(Weekday(_LastDate)) = "Friday" Then
                    _LastDate = _LastDate.AddDays(-2)
                ElseIf WeekdayName(Weekday(_LastDate)) = "Saturday" Then
                    _LastDate = _LastDate.AddDays(-3)

                End If


                _TimeSpan = _LastDate.Subtract(userDate)
                _WeekNo = _TimeSpan.Days / 7

                userDate = userDate.AddDays(+7)
                vcwhere = "T15Sales_Order='" & strSales_Order & "' AND t01bulk='1st Bulk' and T15Year=" & M01.Tables(0).Rows(i)("T15Year") & " and T15Month=" & M01.Tables(0).Rows(i)("T15Month") & ""
                M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "FSTB"), New SqlParameter("@vcWhereClause1", vcwhere))
                If isValidDataset(M02) Then
                    userDate = userDate.AddDays(-14)
                Else
                    userDate = userDate.AddDays(-7)
                End If

                For Z = 1 To _WeekNo
                    Dim culture As System.Globalization.CultureInfo
                    Dim intWeek As Integer
                    Dim _StrWeek As String

                    'If Z = 1 Then
                    '    _weekSt = userDate
                    'Else
                    '    _weekSt = userDate.AddDays(-6)
                    'End If
                    culture = System.Globalization.CultureInfo.CurrentCulture
                    intWeek = culture.Calendar.GetWeekOfYear(userDate, System.Globalization.CalendarWeekRule.FirstFullWeek, DayOfWeek.Thursday)

                    strWeek = "Week " & intWeek

                    ' strWeek = M01.Tables(0).Rows(i)("M43Product_Month")
                    colWork = New DataColumn(strWeek, GetType(String))
                    colWork.MaxLength = 250
                    dataTable.Columns.Add(colWork)

                    userDate = userDate.AddDays(+7)

                Next
                
                i = i + 1
            Next

            '==================================================================


            Return dataTable

            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            con.close()

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)

                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
                con.close()
            End If
        End Try
    End Function

    Function Data_Fill_Projection_Dye(ByVal _Quality_No As String)
        Dim vcwhere As String
        Dim i As Integer
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim M02 As DataSet
        Dim M03 As DataSet

        Dim Value As Double
        Dim _ST As String
        Dim Y As Integer
        Dim strWeek As String
        Dim _Date As Date
        Dim _Code As Integer
        Dim Z As Integer
        Dim _Rowcount As Integer
        Dim _Coloumcount As Integer

        Try

            vcwhere = "select * from P01PARAMETER where P01CODE='PRN'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, vcwhere)
            If isValidDataset(M01) Then
                _Code = M01.Tables(0).Rows(0)("P01NO")
            End If

            _Code = _Code - 1
            _Projection_Code = _Code
            If Microsoft.VisualBasic.Day(Today) > 10 Then
                _Date = Today.AddDays(+30)
                _Date = Month(_Date) & "/1/" & Year(_Date)
                vcwhere = "M43Quality in ('" & _Quality_No & "') and convert(datetime,Ddate,111)>='" & _Date & "'  and M43Count_No=" & _Code & ""
            Else
                _Date = Today
                _Date = Month(_Date) & "/1/" & Year(_Date)
                vcwhere = "M43Quality in ('" & _Quality_No & "') and convert(datetime,Ddate,111)>='" & _Date & "'  and M43Count_No=" & _Code & ""
            End If
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "PRS2"), New SqlParameter("@vcWhereClause1", vcwhere))
            i = 0
            For Each DTRow5 As DataRow In M01.Tables(0).Rows
                If i = 0 Then
                    _DyeMonth = M01.Tables(0).Rows(i)("M43Product_Month")
                    _DyeYear = M01.Tables(0).Rows(i)("M43Year")
                Else
                    _DyeMonth = _DyeMonth & "','" & M01.Tables(0).Rows(i)("M43Product_Month")
                    _DyeYear = _DyeYear & "','" & M01.Tables(0).Rows(i)("M43Year")
                End If
               
                i = i + 1
            Next

            'Filling the data
            vcwhere = "M43Year in ('" & _DyeYear & "') and  M43Product_Month in ('" & _DyeMonth & "') and M43Count_No=" & _Projection_Code & " and m22Quality in ('" & _Quality_No & "')"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "VPR1"), New SqlParameter("@vcWhereClause1", vcwhere))
            i = 0
            For Each DTRow5 As DataRow In M01.Tables(0).Rows
                Dim newRowD1 As DataRow = c_dataCustomer2_Dye.NewRow

                newRowD1("Code") = M01.Tables(0).Rows(i)("Code")
                c_dataCustomer2_Dye.Rows.Add(newRowD1)
                i = i + 1
            Next
            '-------------------------------------------------------------------------
            _Rowcount = 0
            _Coloumcount = 1
            vcwhere = "M43Year in ('" & _DyeYear & "') and  M43Product_Month in ('" & _DyeMonth & "') and M43Count_No=" & _Projection_Code & " and m22Quality in ('" & _Quality_No & "')"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "VPR1"), New SqlParameter("@vcWhereClause1", vcwhere))
            i = 0
            For Each DTRow5 As DataRow In M01.Tables(0).Rows
                _Coloumcount = 1
                If Microsoft.VisualBasic.Day(Today) > 10 Then
                    _Date = Today.AddDays(+30)
                    _Date = Month(_Date) & "/1/" & Year(_Date)
                    vcwhere = "M43Quality in ('" & _Quality_No & "') and convert(datetime,Ddate,111)>='" & _Date & "'  and M43Count_No=" & _Code & ""
                Else
                    _Date = Today
                    _Date = Month(_Date) & "/1/" & Year(_Date)
                    vcwhere = "M43Quality in ('" & _Quality_No & "') and convert(datetime,Ddate,111)>='" & _Date & "'  and M43Count_No=" & _Code & ""
                End If
                M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "PRS2"), New SqlParameter("@vcWhereClause1", vcwhere))
                Z = 0
                For Each DTRow7 As DataRow In M02.Tables(0).Rows
                    vcwhere = "M43Year in ('" & M02.Tables(0).Rows(Z)("M43Year") & "') and  M43Product_Month in ('" & M02.Tables(0).Rows(Z)("M43Product_Month") & "') and M43Count_No=" & _Projection_Code & " and m22Quality in ('" & _Quality_No & "') and Code='" & M01.Tables(0).Rows(i)("Code") & "'"
                    M03 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "VPR"), New SqlParameter("@vcWhereClause1", vcwhere))
                    If isValidDataset(M03) Then
                        dgDye_Projection.Rows(_Rowcount).Cells(_Coloumcount).Value = CInt(M03.Tables(0).Rows(0)("Qty"))
                    End If
                    _Coloumcount = _Coloumcount + 1
                    Z = Z + 1
                Next
                _Rowcount = _Rowcount + 1
                i = i + 1
            Next
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            con.close()

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)

                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
                con.close()
            End If
        End Try
    End Function

    Function Load_Gride_Dye_Grige()
        Dim CustomerDataClass As New frmKnitting_Plan_WithTab
        c_dataCustomer1_Dye = CustomerDataClass.MakeDataTable_Greige_Rady()
        dg_dye_Main.DataSource = c_dataCustomer1_Dye
        With dg_dye_Main
            .DisplayLayout.Bands(0).Columns(1).Width = 80
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 80
            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(3).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(3).Width = 190
            .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Left
            .DisplayLayout.Bands(0).Columns(4).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(4).Width = 90
            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(5).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(5).Width = 90
            .DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            ' .DisplayLayout.Bands(0).Columns(6).AutoEdit = False
            '.DisplayLayout.Bands(0).Columns(6).Width = 90
            '.DisplayLayout.Bands(0).Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
        End With
    End Function

    Function Load_Gride_Dye_Grige_Detailes()
        '2nd Screen
        Dim CustomerDataClass As New frmKnitting_Plan_WithTab
        c_dataCustomer3_Dye = CustomerDataClass.MakeDataTable_Greige_Rady_Details()
        dgDye_Gr.DataSource = c_dataCustomer3_Dye
        With dgDye_Gr
            .DisplayLayout.Bands(0).Columns(1).Width = 80
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 80
            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(3).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(3).Width = 190
            .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Left
            .DisplayLayout.Bands(0).Columns(4).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(4).Width = 90
            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(5).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(5).Width = 90
            .DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            ' .DisplayLayout.Bands(0).Columns(6).AutoEdit = False
            '.DisplayLayout.Bands(0).Columns(6).Width = 90
            ' .DisplayLayout.Bands(0).Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
        End With
    End Function

    Function Load_Gride_Dye_Projection()
        Dim i As Integer

        Dim CustomerDataClass As New frmKnitting_Plan_WithTab
        c_dataCustomer2_Dye = CustomerDataClass.MakeDataTable_Projection_Dye(_DyeQuality)
        dgDye_Projection.DataSource = c_dataCustomer2_Dye
        With dgDye_Projection
            .DisplayLayout.Bands(0).Columns(0).Width = 190

            For i = 1 To .DisplayLayout.Bands(0).Columns.Count - 1
                If .DisplayLayout.Bands(0).Columns.Count = 1 Then

                Else
                    .DisplayLayout.Bands(0).Columns(i).Width = 70
                    .DisplayLayout.Bands(0).Columns(i).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                End If
            Next
        End With
    End Function

    Function Load_Gride_Dye_Capacity_New()
        Dim i As Integer

        Dim CustomerDataClass As New frmKnitting_Plan_WithTab
        c_dataCustomer4_Dye = CustomerDataClass.MakeDataTable_Capacity_Dye(_Dye_LineItem)
        dg_Dye_Detailes.DataSource = c_dataCustomer4_Dye
        With dg_Dye_Detailes
            .DisplayLayout.Bands(0).Columns(0).Width = 190

            For i = 1 To .DisplayLayout.Bands(0).Columns.Count - 1
                If .DisplayLayout.Bands(0).Columns.Count = 1 Then

                Else
                    .DisplayLayout.Bands(0).Columns(i).Width = 70
                    .DisplayLayout.Bands(0).Columns(i).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                End If
            Next
        End With
    End Function

    Private Sub UltraButton22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton22.Click
        If OPR_YDP.Visible = True Then
            OPR_YDP.Visible = False
        Else
            OPR_YDP.Visible = True
        End If
    End Sub

    Private Sub UltraButton17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton17.Click
        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Dim vcWhere As String
        Dim _McBooking_Status As Boolean

        Dim M01 As DataSet
        Dim i As Integer
        Dim ncQryType As String
        Try
            'If CDbl(lblKnt_Balance.Text) > 0 Then
            '    Exit Sub
            'Else

            'End If
            '==============================================
            'Insert Knitting Dash Board
            i = 0
            For Each uRow As UltraGridRow In UltraGrid4.Rows
                If UltraGrid4.Rows(i).Cells(7).Value = True Then
                    _McBooking_Status = True
                    'Calculate Week No

                End If
                i = i + 1
            Next
            '-------------------------------------------------------
            If _McBooking_Status = False Then
                MsgBox("Please allocate the machine", MsgBoxStyle.Information, "Information .....")
                connection.Close()
                Exit Sub
            End If
            i = 0
            For Each uRow As UltraGridRow In UltraGrid1.Rows
                ' MsgBox(UltraGrid1.Rows(i).Cells(5).Value)
                If IsNumeric(UltraGrid1.Rows(i).Cells(6).Value) Then
                    vcWhere = "T12Ref_No=" & Delivary_Ref & " and T12Sales_Order='" & strSales_Order & "' and T12Line_Item=" & strLine_Item & " and T12Stock_Code='" & Trim(UltraGrid1.Rows(i).Cells(3).Value) & "'"
                    M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetUse_Griege_Qty", New SqlParameter("@cQryType", "CGS1"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M01) Then
                    Else
                        ncQryType = "GADD"
                        nvcFieldList1 = "(T12Ref_No," & "T12Sales_Order," & "T12Line_Item," & "T12Date," & "T12Time," & "T12Stock_Code," & "T12Qty," & "T12Status," & "T12Confirm_By) " & "values(" & Delivary_Ref & ",'" & strSales_Order & "'," & strLine_Item & ",'" & Today & "','" & Now & "','" & Trim(UltraGrid1.Rows(i).Cells(3).Value) & "','" & Trim(UltraGrid1.Rows(i).Cells(6).Value) & "','N','-')"
                        up_GetSetConsume_Grige(ncQryType, nvcFieldList1, vcWhere, connection, transaction)
                    End If
                End If


                i = i + 1
            Next

            nvcFieldList1 = "delete from T17Planning_Detailes where T17RefNo=" & Delivary_Ref & " and T17Sales_Order='" & strSales_Order & "' and T17Line_Item=" & strLine_Item & ""
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            nvcFieldList1 = "delete from tmpBlock_KnittingMC where tmpSales_Order='" & strSales_Order & "' and tmpLine_Item=" & strLine_Item & ""
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)


            nvcFieldList1 = "delete from tmp_Knt_Mc where tmpSales_Order='" & strSales_Order & "'"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)


            'T17Planning Data Update
            ncQryType = "IPD"
            nvcFieldList1 = "(T17RefNo," & "T17Date," & "T17Sales_Order," & "T17Line_Item," & "T17Order_Qty," & "T17Ex_FG," & "T17Use_FG," & "T17WIP," & "T17Use_WIP," & "T17Ex_LIB," & "T17Use_LIB," & "T17Balance," & "T17MOQ," & "T17LIB," & "T17Req_Griege," & "T17Req_LIB," & "T17User," & "T17Status) " & "values(" & Delivary_Ref & ",'" & Today & "'," & strSales_Order & ",'" & strLine_Item & "','" & txtOrder_Qty.Text & "','" & txtExcess_FG.Text & "','" & txtUse_FG.Text & "','" & txtWIP.Text & "','" & txtUse_WIP.Text & "','" & txtEx_Lib.Text & "','" & txtEx_LibUse.Text & "','" & CDbl(txtBalance.Text) & "','" & (txtMOQ.Text) & "','" & txtLIB.Text & "','" & (txtReq_Grg.Text) & "','" & txtReg_LIb.Text & "','" & strDisname & "','C')"
            up_GetSettmp_ProjectAllocation(ncQryType, nvcFieldList1, vcWhere, connection, transaction)

            '--------------------------------------------------------------------------------
            nvcFieldList1 = "update tmpYarn_Booking set tmpB_Status='C' where tmpRef='" & Delivary_Ref & "' and tmpB_Status='A'"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            nvcFieldList1 = "update T01Delivary_Request set T01Status='C' where T01Sales_Order='" & strSales_Order & "' and T01Line_Item=" & strLine_Item & " and T01Status='A'"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            transaction.Commit()
            Call Save_Temp_KnittingBord()

            nvcFieldList1 = "T01Sales_Order='" & strSales_Order & "' and T01Status='A'"
            M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "T01D"), New SqlParameter("@vcWhereClause1", nvcFieldList1))
            If isValidDataset(M01) Then
                connection.Close()
                frmDelivaryQuatnew.Load_Gride_SalesOrder()
                frmDelivaryQuatnew.Load_SalesOrder()

                Me.Close()
                Exit Sub
            Else
                UltraTabControl1.Tabs(5).Enabled = True
                UltraTabControl1.SelectedTab = UltraTabControl1.Tabs(5)
            End If

            connection.Close()
            ' Me.Close()
            'frmLoad_Pln.Close()
            frmDelivaryQuatnew.Load_Gride_SalesOrder()
            frmDelivaryQuatnew.Load_SalesOrder()

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)

                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
                connection.Close()
            End If
        End Try
    End Sub

    Private Sub dg_dye_Main_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles dg_dye_Main.Click
        'Dim _Rowindex As Integer
        'Dim _LineItem As String
        'Dim vcWhere As String
        'Dim M01 As DataSet
        'Dim M02 As DataSet
        'Dim M03 As DataSet
        'Dim M04 As DataSet
        'Dim i As Integer
        'Dim con = New SqlConnection()
        'con = DBEngin.GetConnection(True)
        'Dim Value As Double
        'Dim _ST As String
        'Dim Y As Integer
        'Dim strWeek As String
        'Dim Z As Integer
        'Dim _BaseLine_Item As String
        'Dim _BASEQUALITY As String
        'Dim _Trim_Quality As String
        'Dim _Qty As Integer
        'Dim _base30class As String
        'Dim _caractRemove As String
        'Dim _Rcode As String

        'Try
        '    Call Load_Gride_Dye_Grige_Detailes()

        '    _Rowindex = dg_dye_Main.ActiveRow.Index
        '    _LineItem = Trim(dg_dye_Main.Rows(_Rowindex).Cells(1).Text)
        '    _base30class = Trim(dg_dye_Main.Rows(_Rowindex).Cells(2).Text)
        '    _Qty = 0
        '    _caractRemove = "-"
        '    _base30class = (Replace(_base30class, _caractRemove, ""))
        '    _DyeQuality = ""
        '    vcWhere = "T01Sales_Order='" & strSales_Order & "' and T01Line_Item=" & _LineItem & " and T01Status='C'"
        '    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "T01D"), New SqlParameter("@vcWhereClause1", vcWhere))
        '    If isValidDataset(M01) Then
        '        If Trim(M01.Tables(0).Rows(0)("T01Maching")) <> "" Then
        '            _Qty = M01.Tables(0).Rows(0)("T01Qty")
        '            _BaseLine_Item = M01.Tables(0).Rows(0)("T01Maching")
        '            vcWhere = "T01Sales_Order='" & strSales_Order & "' and T01Line_Item=" & Trim(M01.Tables(0).Rows(0)("T01Maching")) & " and T01Status='C'"
        '            M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "T01D"), New SqlParameter("@vcWhereClause1", vcWhere))
        '            If isValidDataset(M02) Then
        '                Value = 0
        '                Dim newRow As DataRow = c_dataCustomer3_Dye.NewRow

        '                newRow("##") = "Body"
        '                newRow("Line Item") = M02.Tables(0).Rows(0)("T01Line_Item")
        '                newRow("Material") = M02.Tables(0).Rows(0)("M01Material_No")
        '                newRow("Description") = M02.Tables(0).Rows(0)("M01Quality")
        '                _Qty = _Qty + M02.Tables(0).Rows(0)("T01Qty")
        '                Value = _Qty
        '                _ST = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
        '                _ST = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
        '                newRow("Quantity") = _ST


        '                Y = 0
        '                vcWhere = "tmpSales_Order='" & strSales_Order & "' and tmpLine_Item=" & M02.Tables(0).Rows(0)("T01Line_Item") & ""
        '                M03 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "SRK"), New SqlParameter("@vcWhereClause1", vcWhere))
        '                For Each DTRow5 As DataRow In M03.Tables(0).Rows

        '                    vcWhere = "tmpSales_Order='" & strSales_Order & "' and tmpLine_Item=" & M02.Tables(0).Rows(0)("T01Line_Item") & " and tmpWeek_No=" & M03.Tables(0).Rows(Y)("tmpWeek_No") & ""
        '                    M04 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "SRK1"), New SqlParameter("@vcWhereClause1", vcWhere))
        '                    If isValidDataset(M04) Then
        '                        strWeek = "Week " & M03.Tables(0).Rows(Y)("tmpWeek_No")
        '                        Value = M04.Tables(0).Rows(0)("Qty")
        '                        _ST = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
        '                        _ST = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
        '                        newRow(strWeek) = _ST

        '                    End If
        '                    Y = Y + 1
        '                Next

        '                c_dataCustomer3_Dye.Rows.Add(newRow)
        '            End If

        '        Else
        '            _BaseLine_Item = _LineItem
        '            vcWhere = "T01Sales_Order='" & strSales_Order & "' and T01Line_Item=" & _LineItem & " and T01Status='C'"
        '            M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "DPB"), New SqlParameter("@vcWhereClause1", vcWhere))
        '            If isValidDataset(M02) Then
        '                Value = 0
        '                Dim newRow As DataRow = c_dataCustomer3_Dye.NewRow
        '                _BASEQUALITY = M02.Tables(0).Rows(0)("M01Quality_No")
        '                newRow("##") = "Body"
        '                newRow("Line Item") = M02.Tables(0).Rows(0)("T01Line_Item")
        '                newRow("Material") = M02.Tables(0).Rows(0)("M01Material_No")
        '                newRow("Description") = M02.Tables(0).Rows(0)("M01Quality")
        '                Value = M02.Tables(0).Rows(0)("T17Req_Griege")
        '                _Qty = _Qty + M02.Tables(0).Rows(0)("T17Req_Griege")
        '                _ST = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
        '                _ST = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
        '                newRow("Quantity") = _ST


        '                Y = 0
        '                vcWhere = "tmpSales_Order='" & strSales_Order & "' and tmpLine_Item=" & Trim(M02.Tables(0).Rows(0)("T01Line_Item")) & ""
        '                M03 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "SRK"), New SqlParameter("@vcWhereClause1", vcWhere))
        '                For Each DTRow5 As DataRow In M03.Tables(0).Rows

        '                    vcWhere = "tmpSales_Order='" & strSales_Order & "' and tmpLine_Item=" & Trim(M02.Tables(0).Rows(0)("T01Line_Item")) & " and tmpWeek_No=" & M03.Tables(0).Rows(Y)("tmpWeek_No") & ""
        '                    M04 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "SRK1"), New SqlParameter("@vcWhereClause1", vcWhere))
        '                    If isValidDataset(M04) Then
        '                        strWeek = "Week " & M03.Tables(0).Rows(Y)("tmpWeek_No")
        '                        Value = M04.Tables(0).Rows(0)("Qty")
        '                        _ST = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
        '                        _ST = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
        '                        newRow(strWeek) = _ST

        '                    End If
        '                    Y = Y + 1
        '                Next

        '                c_dataCustomer3_Dye.Rows.Add(newRow)
        '                '======================================================================
        '                'Load Trim Quality
        '                _Trim_Quality = ""
        '                i = 0
        '                vcWhere = "T01Sales_Order='" & strSales_Order & "' and T01Maching=" & _LineItem & " and T01Status='C'"
        '                M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "DPB"), New SqlParameter("@vcWhereClause1", vcWhere))
        '                For Each DTRow6 As DataRow In M02.Tables(0).Rows
        '                    Value = 0
        '                    Dim newRow3 As DataRow = c_dataCustomer3_Dye.NewRow
        '                    If i = 0 Then
        '                        _Trim_Quality = Trim(M02.Tables(0).Rows(i)("M01Quality_No"))
        '                    Else
        '                        _Trim_Quality = _Trim_Quality & "','" & Trim(M02.Tables(0).Rows(i)("M01Quality_No"))
        '                    End If
        '                    newRow3("##") = "Trim"
        '                    newRow3("Line Item") = M02.Tables(0).Rows(i)("T01Line_Item")
        '                    newRow3("Material") = M02.Tables(0).Rows(i)("M01Material_No")
        '                    newRow3("Description") = M02.Tables(0).Rows(i)("M01Quality")
        '                    Value = M02.Tables(0).Rows(i)("T17Req_Griege")
        '                    _Qty = _Qty + M02.Tables(0).Rows(i)("T17Req_Griege")
        '                    _ST = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
        '                    _ST = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
        '                    newRow3("Quantity") = _ST


        '                    Y = 0
        '                    vcWhere = "tmpSales_Order='" & strSales_Order & "' and tmpLine_Item=" & M02.Tables(0).Rows(0)("T01Line_Item") & ""
        '                    M03 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "SRK"), New SqlParameter("@vcWhereClause1", vcWhere))
        '                    For Each DTRow5 As DataRow In M03.Tables(0).Rows

        '                        vcWhere = "tmpSales_Order='" & strSales_Order & "' and tmpLine_Item=" & M02.Tables(0).Rows(0)("T01Line_Item") & " and tmpWeek_No=" & M03.Tables(0).Rows(Y)("tmpWeek_No") & ""
        '                        M04 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "SRK1"), New SqlParameter("@vcWhereClause1", vcWhere))
        '                        If isValidDataset(M04) Then
        '                            strWeek = "Week " & M03.Tables(0).Rows(Y)("tmpWeek_No")
        '                            Value = M04.Tables(0).Rows(0)("Qty")
        '                            _ST = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
        '                            _ST = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
        '                            newRow3(strWeek) = _ST

        '                        End If
        '                        Y = Y + 1
        '                    Next

        '                    c_dataCustomer3_Dye.Rows.Add(newRow3)
        '                    i = i + 1
        '                Next

        '                Panel16.Visible = False
        '                GroupBox1.Visible = False
        '                dg_dye_Main.Visible = False
        '                '---------------------------------------------------------------------------
        '                'Loading Details NPL/PP/1stBulk
        '                vcWhere = "T01Sales_Order='" & strSales_Order & "' and T01Line_Item='" & _BaseLine_Item & "' and T01Status='C'"
        '                M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "T01D"), New SqlParameter("@vcWhereClause1", vcWhere))
        '                If isValidDataset(M02) Then
        '                    txtDye_Bulk.Text = M02.Tables(0).Rows(0)("T01Bulk")
        '                    txtDye_LD.Text = M02.Tables(0).Rows(0)("T01Lab_Dye")
        '                    If Trim(M02.Tables(0).Rows(0)("T01Lab_Dye")) = "NOT APPROVED" Then
        '                        txtDye_LDApp.Text = M02.Tables(0).Rows(0)("T01POD")
        '                    End If
        '                    txtDye_NPL.Text = M02.Tables(0).Rows(0)("T01NPL")
        '                    If Trim(M02.Tables(0).Rows(0)("T01NPL")) = "NOT APPROVED" Then
        '                        txtDye_NPL_App.Text = M02.Tables(0).Rows(0)("T01NPL_AppDate")
        '                    End If

        '                    txtDye_PP.Text = M02.Tables(0).Rows(0)("T01PP")
        '                    If Trim(M02.Tables(0).Rows(0)("T01PP")) = "NOT APPROVED" Then
        '                        txtDye_PP_App.Text = M02.Tables(0).Rows(0)("T01PP_AppDate")
        '                    End If

        '                End If

        '            End If

        '        End If
        '    End If
        '    '====================================================================================
        '    'CHECK R-CODE

        '    vcWhere = "M16Material='" & _base30class & "'"
        '    M04 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "QRQ"), New SqlParameter("@vcWhereClause1", vcWhere))
        '    If isValidDataset(M04) Then
        '        _Rcode = Trim(M04.Tables(0).Rows(0)("M16R_Code"))
        '        vcWhere = "M14Order='" & Trim(M04.Tables(0).Rows(0)("M16R_Code")) & "'" ' and m14status='CUS APPD'"
        '        M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "RCOD"), New SqlParameter("@vcWhereClause1", vcWhere))
        '        If isValidDataset(M02) Then
        '            'Using Trim Quality
        '            vcWhere = "M49quality='" & _BASEQUALITY & "' and M49trim in ('" & _Trim_Quality & "')"
        '            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "CGLI"), New SqlParameter("@vcWhereClause1", vcWhere))
        '            If isValidDataset(M01) Then
        '                i = 0
        '                lblConstrain.Text = M01.Tables(0).Rows(0)("M49Comment")
        '                i = 0
        '                For Each DTRow5 As DataRow In M01.Tables(0).Rows
        '                    lblConstrain.Text = M01.Tables(0).Rows(0)("M49Comment")
        '                    'MsgBox(Trim(M02.Tables(0).Rows(0)("M14grige")))

        '                    If Trim(M04.Tables(0).Rows(0)("M16Shade_Type")) = "White" Then
        '                        If Trim(M01.Tables(0).Rows(i)("M49Mc_Group")) = "LR" Then
        '                            txtDye_S_LR.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49SR_White")))
        '                            txtDye_D_LR.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49DR_White")))
        '                        ElseIf Trim(M01.Tables(0).Rows(i)("M49Mc_Group")) = "THAN" Then
        '                            txtDye_S_Than.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49SR_White")))
        '                            txtDye_D_Than.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49DR_White")))
        '                        ElseIf Trim(M01.Tables(0).Rows(i)("M49Mc_Group")) = "ECO" Then
        '                            txtDye_S_Eco.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49SR_White")))
        '                            txtDye_D_Eco.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49DR_White")))
        '                        ElseIf Trim(M01.Tables(0).Rows(i)("M49Mc_Group")) = "ECO +" Then
        '                            txtS_Eco1.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49SR_White")))
        '                            txtDye_D_Eco1.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49DR_White")))
        '                        End If
        '                    Else
        '                        If IsDBNull(M02.Tables(0).Rows(0)("M14Criticle")) Then
        '                            If Trim(M01.Tables(0).Rows(i)("M49Mc_Group")) = "LR" Then
        '                                txtDye_S_LR.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49SR_Nomal")))
        '                                txtDye_D_LR.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49DR_Nomal")))
        '                            ElseIf Trim(M01.Tables(0).Rows(i)("M49Mc_Group")) = "THAN" Then
        '                                txtDye_S_Than.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49SR_Nomal")))
        '                                txtDye_D_Than.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49DR_Nomal")))
        '                            ElseIf Trim(M01.Tables(0).Rows(i)("M49Mc_Group")) = "ECO" Then
        '                                txtDye_S_Eco.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49SR_Nomal")))
        '                                txtDye_D_Eco.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49DR_Nomal")))
        '                            ElseIf Trim(M01.Tables(0).Rows(i)("M49Mc_Group")) = "ECO +" Then
        '                                txtS_Eco1.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49SR_Nomal")))
        '                                txtDye_D_Eco1.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49DR_Nomal")))
        '                            End If

        '                        ElseIf Trim(M02.Tables(0).Rows(0)("M14Criticle")) = "Y" Then
        '                            If Trim(M01.Tables(0).Rows(i)("M49Mc_Group")) = "LR" Then
        '                                txtDye_S_LR.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49SR_Critical")))
        '                                txtDye_D_LR.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49DR_Critical")))
        '                            ElseIf Trim(M01.Tables(0).Rows(i)("M49Mc_Group")) = "THAN" Then
        '                                txtDye_S_Than.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49SR_Critical")))
        '                                txtDye_D_Than.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49DR_Critical")))
        '                            ElseIf Trim(M01.Tables(0).Rows(i)("M49Mc_Group")) = "ECO" Then
        '                                txtDye_S_Eco.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49SR_Critical")))
        '                                txtDye_D_Eco.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49DR_Critical")))
        '                            ElseIf Trim(M01.Tables(0).Rows(i)("M49Mc_Group")) = "ECO +" Then
        '                                txtS_Eco1.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49SR_Critical")))
        '                                txtDye_D_Eco1.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49DR_Critical")))
        '                            End If
        '                        ElseIf Trim(M02.Tables(0).Rows(0)("M14Criticle")) = "N" Or Trim(M02.Tables(0).Rows(0)("M14Criticle")) = "" Then
        '                            If Trim(M01.Tables(0).Rows(i)("M49Mc_Group")) = "LR" Then
        '                                txtDye_S_LR.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49SR_Nomal")))
        '                                txtDye_D_LR.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49DR_Nomal")))
        '                            ElseIf Trim(M01.Tables(0).Rows(i)("M49Mc_Group")) = "THAN" Then
        '                                txtDye_S_Than.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49SR_Nomal")))
        '                                txtDye_D_Than.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49DR_Nomal")))
        '                            ElseIf Trim(M01.Tables(0).Rows(i)("M49Mc_Group")) = "ECO" Then
        '                                txtDye_S_Eco.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49SR_Nomal")))
        '                                txtDye_D_Eco.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49DR_Nomal")))
        '                            ElseIf Trim(M01.Tables(0).Rows(i)("M49Mc_Group")) = "ECO +" Then
        '                                txtS_Eco1.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49SR_Nomal")))
        '                                txtDye_D_Eco1.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49DR_Nomal")))
        '                            End If
        '                        End If
        '                    End If
        '                    ' MsgBox(Trim(M02.Tables(0).Rows(0)("M14Criticle")))
        '                    i = i + 1
        '                Next

        '                _DyeQuality = _BASEQUALITY

        '                If Trim(_Trim_Quality) <> "" Then
        '                    _DyeQuality = _DyeQuality & "','" & _Trim_Quality
        '                Else

        '                End If

        '                'Dye Quntity
        '                'Grage Booking
        '                'Developed by Suranga on 2016.7.15
        '                _Qty = 0
        '                vcWhere = "T12Sales_Order='" & strSales_Order & "' "
        '                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "T12Q"), New SqlParameter("@vcWhereClause1", vcWhere))
        '                If isValidDataset(M01) Then
        '                    _Qty = M01.Tables(0).Rows(0)("Qty")
        '                End If

        '                'Knitting Booking
        '                vcWhere = "tmpSales_Order='" & strSales_Order & "' "
        '                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "KPLB"), New SqlParameter("@vcWhereClause1", vcWhere))
        '                If isValidDataset(M01) Then
        '                    _Qty = _Qty + M01.Tables(0).Rows(0)("Qty")
        '                End If

        '                lblDye_Qty.Text = (_Qty.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
        '                lblDye_Qty.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", _Qty))

        '                txtDye_S_LR.Appearance.BackColor = Color.White
        '                txtDye_D_LR.Appearance.BackColor = Color.White
        '                txtDye_S_Than.Appearance.BackColor = Color.White
        '                txtDye_D_Than.Appearance.BackColor = Color.White
        '                txtS_Eco1.Appearance.BackColor = Color.White
        '                txtDye_D_Eco1.Appearance.BackColor = Color.White
        '                txtDye_S_Eco.Appearance.BackColor = Color.White
        '                txtDye_D_Eco.Appearance.BackColor = Color.White
        '                lblPrevious_St_Code.Text = "-"
        '                'PREVIOUS DYE MACHINE GROUP
        '                vcWhere = "M16R_CODE='" & _Rcode & "' "
        '                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "DPBT"), New SqlParameter("@vcWhereClause1", vcWhere))
        '                If isValidDataset(M01) Then
        '                    ' MsgBox(UCase(Trim(M01.Tables(0).Rows(0)("tmpGroup"))))
        '                    lblPrevious_St_Code.Text = Trim(M01.Tables(0).Rows(0)("tmpStock_Code"))
        '                    If UCase(Trim(M01.Tables(0).Rows(0)("tmpGroup"))) = "LR" Then
        '                        txtDye_S_LR.Appearance.BackColor = Color.Yellow
        '                        txtDye_D_LR.Appearance.BackColor = Color.Yellow
        '                    ElseIf UCase(Trim(M01.Tables(0).Rows(0)("tmpGroup"))) = "ECO" Then
        '                        txtDye_S_Eco.Appearance.BackColor = Color.Yellow
        '                        txtDye_D_Eco.Appearance.BackColor = Color.Yellow
        '                    ElseIf UCase(Trim(M01.Tables(0).Rows(0)("tmpGroup"))) = "THAN" Then
        '                        txtDye_S_Than.Appearance.BackColor = Color.Yellow
        '                        txtDye_D_Than.Appearance.BackColor = Color.Yellow
        '                    ElseIf UCase(Trim(M01.Tables(0).Rows(0)("tmpGroup"))) = "ECO +" Then
        '                        txtS_Eco1.Appearance.BackColor = Color.Yellow
        '                        txtDye_D_Eco1.Appearance.BackColor = Color.Yellow
        '                    End If
        '                    'End If
        '                End If
        '                DBEngin.CloseConnection(con)
        '                con.ConnectionString = ""
        '                con.close()

        '                Call Load_Gride_Dye_Projection()
        '                Call Data_Fill_Projection_Dye(_DyeQuality)
        '            Else
        '                'If Trim(M04.Tables(0).Rows(0)("m16shade_type")) = "Marls" Or Trim(M04.Tables(0).Rows(0)("m16shade_type")) = "Yarn Dyes" Then
        '                '    txtDye_S_LR.Text = "200"
        '                '    txtDye_S_Than.Text = "200"
        '                '    txtDye_S_Eco.Text = "200"
        '                '    txtDye_D_Eco.Text = "200"
        '                '    txtDye_D_Eco1.Text = "200"
        '                '    txtDye_D_LR.Text = "200"
        '                '    txtDye_D_Than.Text = "200"
        '                '    txtS_Eco1.Text = "200"

        '                'End If
        '                'Using Base Quality
        '                vcWhere = "M49quality='" & _BASEQUALITY & "'"
        '                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "CGLI"), New SqlParameter("@vcWhereClause1", vcWhere))
        '                i = 0
        '                For Each DTRow5 As DataRow In M01.Tables(0).Rows
        '                    lblConstrain.Text = M01.Tables(0).Rows(0)("M49Comment")
        '                    'MsgBox(Trim(M02.Tables(0).Rows(0)("M14grige")))
        '                    If Trim(M02.Tables(0).Rows(0)("M14Criticle")) = "Y" And Trim(M02.Tables(0).Rows(0)("M14grige")) = "L" Then
        '                        If Trim(M01.Tables(0).Rows(i)("M49Mc_Group")) = "LR" Then
        '                            txtDye_S_LR.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49SR_Nomal")))
        '                        ElseIf Trim(M01.Tables(0).Rows(i)("M49Mc_Group")) = "THAN" Then
        '                            txtDye_S_Than.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49SR_Nomal")))
        '                        ElseIf Trim(M01.Tables(0).Rows(i)("M49Mc_Group")) = "ECO" Then
        '                            txtDye_S_Eco.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49SR_Nomal")))
        '                        ElseIf Trim(M01.Tables(0).Rows(i)("M49Mc_Group")) = "ECO+" Then
        '                            txtS_Eco1.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49SR_Nomal")))
        '                        End If
        '                    ElseIf Trim(M02.Tables(0).Rows(0)("M14Criticle")) = "Y" And Trim(M02.Tables(0).Rows(0)("M14grige")) = "D" Then
        '                        If Trim(M01.Tables(0).Rows(i)("M49Mc_Group")) = "LR" Then
        '                            txtDye_S_LR.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49SR_Critical")))
        '                        ElseIf Trim(M01.Tables(0).Rows(i)("M49Mc_Group")) = "THAN" Then
        '                            txtDye_S_Than.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49SR_Critical")))
        '                        ElseIf Trim(M01.Tables(0).Rows(i)("M49Mc_Group")) = "ECO" Then
        '                            txtDye_S_Eco.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49SR_Critical")))
        '                        ElseIf Trim(M01.Tables(0).Rows(i)("M49Mc_Group")) = "ECO+" Then
        '                            txtS_Eco1.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49SR_Critical")))
        '                        End If
        '                    ElseIf Trim(M02.Tables(0).Rows(0)("M14Criticle")) = "N" And Trim(M02.Tables(0).Rows(0)("M14grige")) = "D" Then
        '                        If Trim(M01.Tables(0).Rows(i)("M49Mc_Group")) = "LR" Then
        '                            txtDye_D_LR.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49DR_Critical")))
        '                        ElseIf Trim(M01.Tables(0).Rows(i)("M49Mc_Group")) = "THAN" Then
        '                            txtDye_D_Than.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49DR_Critical")))
        '                        ElseIf Trim(M01.Tables(0).Rows(i)("M49Mc_Group")) = "ECO" Then
        '                            txtDye_D_Eco.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49DR_Critical")))
        '                        ElseIf Trim(M01.Tables(0).Rows(i)("M49Mc_Group")) = "ECO+" Then
        '                            txtDye_D_Eco1.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49DR_Critical")))
        '                        End If
        '                    ElseIf Trim(M02.Tables(0).Rows(0)("M14Criticle")) = "N" And Trim(M02.Tables(0).Rows(0)("M14grige")) = "L" Then
        '                        If Trim(M01.Tables(0).Rows(i)("M49Mc_Group")) = "LR" Then
        '                            txtDye_D_LR.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49DR_Nomal")))
        '                        ElseIf Trim(M01.Tables(0).Rows(i)("M49Mc_Group")) = "THAN" Then
        '                            txtDye_D_Than.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49DR_Nomal")))
        '                        ElseIf Trim(M01.Tables(0).Rows(i)("M49Mc_Group")) = "ECO" Then
        '                            txtDye_D_Eco.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49DR_Nomal")))
        '                        ElseIf Trim(M01.Tables(0).Rows(i)("M49Mc_Group")) = "ECO+" Then
        '                            txtDye_D_Eco1.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49DR_Nomal")))
        '                        End If
        '                    End If
        '                    i = i + 1
        '                Next
        '                '======================================================================================
        '                'Projection 
        '                _DyeQuality = _BASEQUALITY

        '                If _Trim_Quality <> "" Then

        '                    _DyeQuality = _DyeQuality & "','" & _Trim_Quality
        '                End If


        '                lblDye_Qty.Text = (_Qty.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
        '                lblDye_Qty.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", _Qty))

        '                DBEngin.CloseConnection(con)
        '                con.ConnectionString = ""
        '                con.close()

        '                Call Load_Gride_Dye_Projection()
        '                Call Data_Fill_Projection_Dye(_DyeQuality)
        '            End If
        '        Else
        '            MsgBox("Can't Find the R-Code.Please Inform to the Merchant", MsgBoxStyle.Information, "Technova ......")
        '            DBEngin.CloseConnection(con)
        '            con.ConnectionString = ""
        '            con.close()
        '        End If

        '    End If


        'Catch returnMessage As EvaluateException
        '    If returnMessage.Message <> Nothing Then
        '        MessageBox.Show(returnMessage.Message)

        '        DBEngin.CloseConnection(con)
        '        con.ConnectionString = ""
        '        con.close()
        '    End If
        'End Try
    End Sub

    Private Sub UltraButton23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton23.Click
        Panel16.Visible = True
        GroupBox1.Visible = True
        dg_dye_Main.Visible = True

        Panel17.Visible = False
        OPR_Dye.Visible = False

    End Sub

    Function Dye_Cren2()
        Dim I As Integer
        I = 0
        '  MsgBox(dgDye_Gr.Rows.Count)
        For Each uRow As UltraGridRow In dgDye_Gr.Rows
            If I = 0 Then

                _Dye_LineItem = Trim(dgDye_Gr.Rows(I).Cells(1).Text)
            Else
                _Dye_LineItem = _Dye_LineItem & "','" & Trim(dgDye_Gr.Rows(I).Cells(1).Text)
            End If
            I = I + 1
        Next
        ' Call BindUltraDropDown1()
        Call Load_Gride_Dye_Capacity_New() 'Dye Capacity

        For I = 0 To 5
            Dim newRow3 As DataRow = c_dataCustomer4_Dye.NewRow
            If I = 0 Then
                newRow3("##") = "Overroll Capacity(Kg)"
            ElseIf I = 1 Then
                newRow3("##") = "Open Capacity(Kg)"
            ElseIf I = 2 Then
                newRow3("##") = "Previous MC Group Capacity(Kg)"
            ElseIf I = 3 Then
                newRow3("##") = "Open Capacity Previous MC Group(Kg)"
            ElseIf I = 4 Then
                newRow3("##") = "Flow Plan"
            ElseIf I = 5 Then
                newRow3("##") = "Fix Plan"
            End If
            c_dataCustomer4_Dye.Rows.Add(newRow3)
        Next


        'dg_Dye_Detailes.Rows(1).Cells(1).Activation = Activation.NoEdit

        'dg_Dye_Detailes.Rows(4).Cells(2).ValueList = Me.UltraDropDown4
        'Dim newRow4 As DataRow = c_dataCustomer4_Dye.NewRow
        'newRow3("##") = "Available Capacity"
        'c_dataCustomer4_Dye.Rows.Add(newRow4)
        For I = 1 To dg_Dye_Detailes.DisplayLayout.Bands(0).Columns.Count
            Dim _St As String
            _St = (dg_Dye_Detailes.DisplayLayout.Bands(0).Columns(0).Header.Caption)
            dg_Dye_Detailes.Rows(0).Cells(I).Activation = Activation.NoEdit
            dg_Dye_Detailes.Rows(1).Cells(I).Activation = Activation.NoEdit
            dg_Dye_Detailes.Rows(2).Cells(I).Activation = Activation.NoEdit
            dg_Dye_Detailes.Rows(3).Cells(I).Activation = Activation.NoEdit

            Call BindUltraDropDown1()
            dg_Dye_Detailes.Rows(4).Cells(I).ValueList = Me.UltraDropDown4
            dg_Dye_Detailes.Rows(5).Cells(I).Style = ColumnStyle.CheckBox
            dg_Dye_Detailes.Rows(5).Cells(I).Value = False
            'Dim checkColumn As UltraGridRow = dg_Dye_Detailes.DisplayLayout.Bands(0).Columns.Add(_St, _St)
            'checkColumn.DataType = GetType(Boolean)
            'checkColumn.CellActivation = Activation.AllowEdit
            'checkColumn.Header.VisiblePosition = 0

        Next

        Call Load_Gride_Dye_Capacity()

    End Function
    Private Sub UltraButton24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton24.Click
        On Error Resume Next
        Dim I As Integer

        If Panel17.Visible = False Then
            Panel16.Visible = False
            GroupBox1.Visible = False
            dg_dye_Main.Visible = False

            Panel17.Visible = True
            OPR_Dye.Visible = True

            I = 0
            '  MsgBox(dgDye_Gr.Rows.Count)
            For Each uRow As UltraGridRow In dgDye_Gr.Rows
                If I = 0 Then

                    _Dye_LineItem = Trim(dgDye_Gr.Rows(I).Cells(1).Text)
                Else
                    _Dye_LineItem = _Dye_LineItem & "','" & Trim(dgDye_Gr.Rows(I).Cells(1).Text)
                End If
                I = I + 1
            Next
            ' Call BindUltraDropDown1()
            Call Load_Gride_Dye_Capacity_New() 'Dye Capacity

            For I = 0 To 5
                Dim newRow3 As DataRow = c_dataCustomer4_Dye.NewRow
                If I = 0 Then
                    newRow3("##") = "Overroll Plant Filling(Kg)"
                ElseIf I = 1 Then
                    newRow3("##") = "Open Capacity(Kg)"
                ElseIf I = 2 Then
                    newRow3("##") = "Selected MC Group Filling(Kg)"
                ElseIf I = 3 Then
                    newRow3("##") = "Open Capacity Selected MC Group(Kg)"
                ElseIf I = 4 Then
                    newRow3("##") = "Flow Plan"
                ElseIf I = 5 Then
                    newRow3("##") = "Fix Plan"
                End If
                c_dataCustomer4_Dye.Rows.Add(newRow3)
            Next


            'dg_Dye_Detailes.Rows(1).Cells(1).Activation = Activation.NoEdit

            'dg_Dye_Detailes.Rows(4).Cells(2).ValueList = Me.UltraDropDown4
            'Dim newRow4 As DataRow = c_dataCustomer4_Dye.NewRow
            'newRow3("##") = "Available Capacity"
            'c_dataCustomer4_Dye.Rows.Add(newRow4)
            ' MsgBox(dg_Dye_Detailes.DisplayLayout.Bands(0).Columns.Count)
            For I = 1 To dg_Dye_Detailes.DisplayLayout.Bands(0).Columns.Count
                Dim _St As String
                _St = (dg_Dye_Detailes.DisplayLayout.Bands(0).Columns(0).Header.Caption)
                dg_Dye_Detailes.Rows(0).Cells(I).Activation = Activation.NoEdit
                dg_Dye_Detailes.Rows(1).Cells(I).Activation = Activation.NoEdit
                dg_Dye_Detailes.Rows(2).Cells(I).Activation = Activation.NoEdit
                dg_Dye_Detailes.Rows(3).Cells(I).Activation = Activation.NoEdit

                Call BindUltraDropDown1()
                dg_Dye_Detailes.Rows(4).Cells(I).ValueList = Me.UltraDropDown4
                dg_Dye_Detailes.Rows(5).Cells(I).Style = ColumnStyle.CheckBox
                dg_Dye_Detailes.Rows(5).Cells(I).Value = False
                'Dim checkColumn As UltraGridRow = dg_Dye_Detailes.DisplayLayout.Bands(0).Columns.Add(_St, _St)
                'checkColumn.DataType = GetType(Boolean)
                'checkColumn.CellActivation = Activation.AllowEdit
                'checkColumn.Header.VisiblePosition = 0

            Next

            Call Load_Gride_Dye_Capacity()

            chkE1.Checked = False
            chkE2.Checked = False
            chkLR1.Checked = False
            chkLR2.Checked = False
            chkEco1.Checked = False
            chkEco2.Checked = False
            chkThan1.Checked = False
            chkTHAN2.Checked = False

        Else
            Panel17.Visible = False
            OPR_Dye.Visible = False

        End If
    End Sub

    Function Load_Gride_Dye_Capacity()

        Dim vcwhere As String
        Dim i As Integer
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim M02 As DataSet
        Dim M03 As DataSet
        Dim T01 As DataSet
        Dim T02 As DataSet

        Dim Value As Double
        Dim _ST As String
        Dim Y As Integer
        Dim strWeek As String
        Dim _Date As Date
        Dim _Code As Integer
        Dim Z As Integer
        Dim _WeekNo As Integer
        Dim _LaneH As Integer
        Dim Y1 As Integer
        Dim _MCNo As Integer

        Dim _columcount As Integer

        Try

            '' add CustomerID column to key array and bind to DataTable
            ' Dim Keys(0) As DataColumn
            vcwhere = "select * from P01PARAMETER where P01CODE='PRN'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, vcwhere)
            If isValidDataset(M01) Then
                _Code = M01.Tables(0).Rows(0)("P01NO")
            End If

            _Code = _Code - 1



            _columcount = 1

            ' colWork.ReadOnly = True

            i = 0
            vcwhere = "T15Sales_Order='" & strSales_Order & "' and T15Line_Item in ('" & _Dye_LineItem & "') "
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "PROX"), New SqlParameter("@vcWhereClause1", vcwhere))
            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                Dim remain As Integer
                Dim noOfWeek As Integer
                Dim userDate As Date
                Dim _LastDate As Date
                Dim _TimeSpan As TimeSpan

                userDate = DateTime.Parse(M01.Tables(0).Rows(i)("T15Month") & "/1/" & M01.Tables(0).Rows(i)("T15Year"))
                ' MsgBox(WeekdayName(Weekday(userDate)))
                If WeekdayName(Weekday(userDate)) = "Sunday" Then
                    userDate = userDate.AddDays(-3)
                ElseIf WeekdayName(Weekday(userDate)) = "Monday" Then
                    userDate = userDate.AddDays(-4)
                ElseIf WeekdayName(Weekday(userDate)) = "Tuesday" Then
                    userDate = userDate.AddDays(-5)
                ElseIf WeekdayName(Weekday(userDate)) = "Thusday" Then
                    'userDate = userDate.AddDays(-1)
                ElseIf WeekdayName(Weekday(userDate)) = "Friday" Then
                    userDate = userDate.AddDays(-1)
                ElseIf WeekdayName(Weekday(userDate)) = "Saturday" Then
                    userDate = userDate.AddDays(-2)

                End If

                _LastDate = DateTime.Parse(M01.Tables(0).Rows(i)("T15Month") & "/1/" & M01.Tables(0).Rows(i)("T15Year"))
                ' MsgBox(Date.DaysInMonth(_LastDate.Year, _LastDate.Month))
                _LastDate = DateTime.Parse(M01.Tables(0).Rows(i)("T15Month") & "/" & Date.DaysInMonth(_LastDate.Year, _LastDate.Month) & "/" & M01.Tables(0).Rows(i)("T15Year"))

                If WeekdayName(Weekday(_LastDate)) = "Sunday" Then
                    _LastDate = _LastDate.AddDays(-4)
                ElseIf WeekdayName(Weekday(_LastDate)) = "Monday" Then
                    _LastDate = _LastDate.AddDays(-5)
                ElseIf WeekdayName(Weekday(_LastDate)) = "Tuesday" Then
                    _LastDate = _LastDate.AddDays(-6)
                ElseIf WeekdayName(Weekday(_LastDate)) = "Thusday" Then
                    'userDate = userDate.AddDays(-1)
                ElseIf WeekdayName(Weekday(_LastDate)) = "Friday" Then
                    _LastDate = _LastDate.AddDays(-2)
                ElseIf WeekdayName(Weekday(_LastDate)) = "Saturday" Then
                    _LastDate = _LastDate.AddDays(-3)

                End If


                _TimeSpan = _LastDate.Subtract(userDate)
                _WeekNo = _TimeSpan.Days / 7

                userDate = userDate.AddDays(+7)
                vcwhere = "T15Sales_Order='" & strSales_Order & "' AND t01bulk='1st Bulk' and T15Year=" & M01.Tables(0).Rows(i)("T15Year") & " and T15Month=" & M01.Tables(0).Rows(i)("T15Month") & ""
                M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "FSTB"), New SqlParameter("@vcWhereClause1", vcwhere))
                If isValidDataset(M02) Then
                    userDate = userDate.AddDays(-14)
                Else
                    userDate = userDate.AddDays(-7)
                End If

                For Z = 1 To _WeekNo
                    Dim culture As System.Globalization.CultureInfo
                    Dim intWeek As Integer
                    Dim _StrWeek As String
                    Dim _StartTime As Date
                    Dim _EndTime As Date
                    Dim _CapacityHR As Integer
                    Dim _UserCapacityHR As Integer
                    Dim _OpencapacityHR As Integer
                    Dim _Baseshade As String
                    Dim _BaseMaterial As String
                    Dim characterToRemove As String
                    Dim _Operning_Capacity As Double
                    Dim _WeekStart As Date

                    _StartTime = userDate & " " & "7:30AM"
                    _EndTime = userDate.AddDays(+7)

                    _EndTime = _EndTime & " " & "7:30AM"

                    If Z = 1 Then
                        _WeekStart = userDate
                    Else

                        _WeekStart = userDate.AddDays(-6)
                    End If
                    culture = System.Globalization.CultureInfo.CurrentCulture
                    intWeek = culture.Calendar.GetWeekOfYear(userDate, System.Globalization.CalendarWeekRule.FirstFullWeek, DayOfWeek.Thursday)

                    vcwhere = "tmpYear=" & Year(userDate) & " and tmpWeek=" & intWeek & ""
                    dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "DPB2"), New SqlParameter("@vcWhereClause1", vcwhere))
                    If isValidDataset(dsUser) Then

                        dg_Dye_Detailes.Rows(0).Cells(_columcount).Value = CInt(dsUser.Tables(0).Rows(0)("Qty"))
                    End If
                    'CALCULATE FREE CAPACITY HR

                    _CapacityHR = 24 * 7 * 0.92

                    characterToRemove = "-"
                    '_BaseMaterial = dgDye_Gr.Rows(0).Cells(2).Text
                    '_BaseMaterial = (Replace(_BaseMaterial, characterToRemove, ""))

                    vcwhere = "m16Quality in ('" & _DyeQuality & "') and m16material='" & _base30class & "'"
                    dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "QRCD"), New SqlParameter("@vcWhereClause1", vcwhere))
                    If isValidDataset(dsUser) Then
                        _Baseshade = Trim(dsUser.Tables(0).Rows(0)("M16Shade_Type"))
                    End If
                    If _Baseshade = "white" Then
                        _CapacityHR = CInt(_CapacityHR / 6)
                    ElseIf _Baseshade = "Marls" Then
                        _CapacityHR = CInt(_CapacityHR / 5)
                    ElseIf _Baseshade = "Yarn Dyes" Then
                        _CapacityHR = CInt(_CapacityHR / 5)
                    Else
                        _CapacityHR = CInt(_CapacityHR / 12)
                    End If

                    _Operning_Capacity = 0
                    ' Z = 0
                    'COMMENT REQUIED BY LALITH ON 2016.4.23 

                    'vcwhere = "M14quality='" & _DyeQuality & "' " 'and m14status='CUS APPD'"
                    'M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "RCOD"), New SqlParameter("@vcWhereClause1", vcwhere))
                    'If isValidDataset(M02) Then
                    '    vcwhere = "M49quality='" & _DyeQuality & "'"
                    '    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "CGLI"), New SqlParameter("@vcWhereClause1", vcwhere))
                    '    Z = 0
                    '    If isValidDataset(M01) Then
                    '        For Each DTRow5 As DataRow In M01.Tables(0).Rows

                    '            Y1 = 0

                    '            vcwhere = "M50MC_Group='" & Trim(M01.Tables(0).Rows(Z)("M49MC_Group")) & "'"
                    '            dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "LAN"), New SqlParameter("@vcWhereClause1", vcwhere))
                    '            For Each DTRow6 As DataRow In dsUser.Tables(0).Rows
                    '                _LaneH = dsUser.Tables(0).Rows(Y1)("Qty")
                    '                _MCNo = dsUser.Tables(0).Rows(Y1)("MCNo")

                    '                If Trim(M02.Tables(0).Rows(0)("M14Criticle")) = "Y" And Trim(M02.Tables(0).Rows(0)("M14grige")) = "L" Then
                    '                    If Trim(M01.Tables(0).Rows(Z)("M49DR_Nomal")) > 0 Then
                    '                        _Operning_Capacity = (CInt(Trim(M01.Tables(0).Rows(Z)("M49DR_Nomal"))) * _CapacityHR * _LaneH * _MCNo) + _Operning_Capacity
                    '                    Else
                    '                        _Operning_Capacity = (CInt(Trim(M01.Tables(0).Rows(Z)("M49SR_Nomal"))) * _CapacityHR * _LaneH * _MCNo) + _Operning_Capacity
                    '                    End If

                    '                ElseIf Trim(M02.Tables(0).Rows(0)("M14Criticle")) = "Y" And Trim(M02.Tables(0).Rows(0)("M14grige")) = "D" Then

                    '                    If Trim(M01.Tables(0).Rows(Z)("M49DR_Critical")) > 0 Then
                    '                        _Operning_Capacity = (CInt(Trim(M01.Tables(0).Rows(Z)("M49DR_Critical"))) * _CapacityHR * _LaneH * _MCNo) + _Operning_Capacity
                    '                    Else
                    '                        _Operning_Capacity = (CInt(Trim(M01.Tables(0).Rows(Z)("M49SR_Critical"))) * _CapacityHR * _LaneH * _MCNo) + _Operning_Capacity
                    '                    End If

                    '                ElseIf Trim(M02.Tables(0).Rows(0)("M14Criticle")) = "N" And Trim(M02.Tables(0).Rows(0)("M14grige")) = "D" Then

                    '                    If Trim(M01.Tables(0).Rows(Z)("M49DR_Critical")) > 0 Then
                    '                        _Operning_Capacity = (CInt(Trim(M01.Tables(0).Rows(Z)("M49DR_Critical"))) * _CapacityHR * _LaneH * _MCNo) + _Operning_Capacity
                    '                    Else
                    '                        _Operning_Capacity = (CInt(Trim(M01.Tables(0).Rows(Z)("M49SR_Critical"))) * _CapacityHR * _LaneH * _MCNo) + _Operning_Capacity
                    '                    End If

                    '                ElseIf Trim(M02.Tables(0).Rows(0)("M14Criticle")) = "N" And Trim(M02.Tables(0).Rows(0)("M14grige")) = "L" Then
                    '                    If Trim(M01.Tables(0).Rows(Z)("M49DR_Nomal")) > 0 Then
                    '                        _Operning_Capacity = (CInt(Trim(M01.Tables(0).Rows(Z)("M49DR_Nomal"))) * _CapacityHR * _LaneH * _MCNo) + _Operning_Capacity
                    '                    Else
                    '                        _Operning_Capacity = (CInt(Trim(M01.Tables(0).Rows(Z)("M49SR_Nomal"))) * _CapacityHR * _LaneH * _MCNo) + _Operning_Capacity
                    '                    End If
                    '                ElseIf Trim(M02.Tables(0).Rows(0)("M14Criticle")) = "" Then
                    '                    If Trim(M01.Tables(0).Rows(Z)("M49DR_Nomal")) > 0 Then
                    '                        _Operning_Capacity = (CInt(Trim(M01.Tables(0).Rows(Z)("M49DR_Nomal"))) * _CapacityHR * _LaneH * _MCNo) + _Operning_Capacity
                    '                    Else
                    '                        _Operning_Capacity = (CInt(Trim(M01.Tables(0).Rows(Z)("M49SR_Nomal"))) * _CapacityHR * _LaneH * _MCNo) + _Operning_Capacity
                    '                    End If
                    '                End If
                    '                Y1 = Y1 + 1
                    '            Next

                    '            Z = Z + 1
                    '        Next
                    '    Else

                    '        ' vcwhere = "M50MC_Group='" & Trim(M01.Tables(0).Rows(Z)("M49MC_Group")) & "'"
                    '        dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "LAN1"), New SqlParameter("@vcWhereClause1", vcwhere))
                    '        If isValidDataset(dsUser) Then
                    '            _LaneH = dsUser.Tables(0).Rows(0)("Qty")
                    '            _MCNo = dsUser.Tables(0).Rows(0)("MCNo")
                    '        End If
                    '        _Operning_Capacity = (200 * _CapacityHR * _LaneH * _MCNo) + _Operning_Capacity
                    '        End If


                    'End If

                    'OVAROLL PLANT CAPACITY 233TNS - LALITH
                    'Z = 0
                    'dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "LAN1"), New SqlParameter("@vcWhereClause1", vcwhere))
                    'For Each DTRow6 As DataRow In dsUser.Tables(0).Rows
                    '    _LaneH = dsUser.Tables(0).Rows(Z)("Qty")
                    '    _MCNo = dsUser.Tables(0).Rows(Z)("MCNo")

                    '    _Operning_Capacity = (233 * _CapacityHR * _LaneH * _MCNo) + _Operning_Capacity
                    '    Z = Z + 1
                    'Next
                    _Operning_Capacity = "233,000"
                    If dg_Dye_Detailes.Rows(0).Cells(_columcount).Text <> "" Then
                        dg_Dye_Detailes.Rows(1).Cells(_columcount).Value = CInt(_Operning_Capacity) - dg_Dye_Detailes.Rows(0).Cells(_columcount).Value
                    Else
                        dg_Dye_Detailes.Rows(1).Cells(_columcount).Value = CInt(_Operning_Capacity)
                    End If
                    '=====================================================================================
                    'PREVIOUS MACHINE GROUP
                    If txtDye_Bulk.Text = "REPEAT" Then
                        vcwhere = "tmp30_Class='" & _BaseMaterial & "'"
                        M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "DPMG"), New SqlParameter("@vcWhereClause1", vcwhere))
                        If isValidDataset(M02) Then
                            vcwhere = "tmpYear=" & Year(userDate) & " and tmpWeek=" & intWeek & " AND tmpGroup='" & Trim(M02.Tables(0).Rows(0)("tmpGroup")) & "'"
                            dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "DPB2"), New SqlParameter("@vcWhereClause1", vcwhere))
                            If isValidDataset(dsUser) Then
                                dg_Dye_Detailes.Rows(2).Cells(_columcount).Value = CInt(dsUser.Tables(0).Rows(0)("Qty"))

                                _CapacityHR = 24 * 7 * 0.92

                                vcwhere = "m16Quality='" & _DyeQuality & "' and m16material='" & _BaseMaterial & "'"
                                M03 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "QRCD"), New SqlParameter("@vcWhereClause1", vcwhere))
                                If isValidDataset(M03) Then
                                    _Baseshade = Trim(M03.Tables(0).Rows(0)("M16Shade_Type"))
                                End If
                                If _Baseshade = "white" Then
                                    _CapacityHR = CInt(_CapacityHR / 6)
                                ElseIf _Baseshade = "Marls" Then
                                    _CapacityHR = CInt(_CapacityHR / 5)
                                ElseIf _Baseshade = "Yarn Dyes" Then
                                    _CapacityHR = CInt(_CapacityHR / 5)
                                Else
                                    _CapacityHR = CInt(_CapacityHR / 12)
                                End If

                                _Operning_Capacity = 0
                                Z = 0
                                vcwhere = "M14quality='" & _DyeQuality & "'" ' and m14status='CUS APPD'"
                                M03 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "RCOD"), New SqlParameter("@vcWhereClause1", vcwhere))
                                If isValidDataset(M03) Then
                                    vcwhere = "M49quality='" & _DyeQuality & "' AND M49MC_GROUP='" & Trim(M02.Tables(0).Rows(0)("tmpGroup")) & "'"
                                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "CGLI"), New SqlParameter("@vcWhereClause1", vcwhere))
                                    Z = 0
                                    For Each DTRow5 As DataRow In M01.Tables(0).Rows
                                        'Dim _LaneH As Integer
                                        'Dim Y1 As Integer
                                        'Dim _MCNo As Integer
                                        Y1 = 0

                                        vcwhere = "M50MC_Group='" & Trim(M01.Tables(0).Rows(Z)("M49MC_Group")) & "'"
                                        T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "LAN"), New SqlParameter("@vcWhereClause1", vcwhere))
                                        For Each DTRow6 As DataRow In T01.Tables(0).Rows
                                            _LaneH = T01.Tables(0).Rows(Y1)("Qty")
                                            _MCNo = T01.Tables(0).Rows(Y1)("MCNO")


                                            If Trim(M03.Tables(0).Rows(0)("M14Criticle")) = "Y" And Trim(M03.Tables(0).Rows(0)("M14grige")) = "L" Then
                                                If Trim(M01.Tables(0).Rows(Z)("M49DR_Nomal")) > 0 Then
                                                    _Operning_Capacity = (CInt(Trim(M01.Tables(0).Rows(Z)("M49DR_Nomal"))) * _CapacityHR * _LaneH * _MCNo) + _Operning_Capacity
                                                Else
                                                    _Operning_Capacity = (CInt(Trim(M01.Tables(0).Rows(Z)("M49SR_Nomal"))) * _CapacityHR * _LaneH * _MCNo) + _Operning_Capacity
                                                End If

                                            ElseIf Trim(M03.Tables(0).Rows(0)("M14Criticle")) = "Y" And Trim(M03.Tables(0).Rows(0)("M14grige")) = "D" Then

                                                If Trim(M01.Tables(0).Rows(Z)("M49DR_Critical")) > 0 Then
                                                    _Operning_Capacity = (CInt(Trim(M01.Tables(0).Rows(Z)("M49DR_Critical"))) * _CapacityHR * _LaneH * _MCNo) + _Operning_Capacity
                                                Else
                                                    _Operning_Capacity = (CInt(Trim(M01.Tables(0).Rows(Z)("M49SR_Critical"))) * _CapacityHR * _LaneH * _MCNo) + _Operning_Capacity
                                                End If

                                            ElseIf Trim(M03.Tables(0).Rows(0)("M14Criticle")) = "N" And Trim(M03.Tables(0).Rows(0)("M14grige")) = "D" Then

                                                If Trim(M01.Tables(0).Rows(Z)("M49DR_Critical")) > 0 Then
                                                    _Operning_Capacity = (CInt(Trim(M01.Tables(0).Rows(Z)("M49DR_Critical"))) * _CapacityHR * _LaneH * _MCNo) + _Operning_Capacity
                                                Else
                                                    _Operning_Capacity = (CInt(Trim(M01.Tables(0).Rows(Z)("M49SR_Critical"))) * _CapacityHR * _LaneH * _MCNo) + _Operning_Capacity
                                                End If

                                            ElseIf Trim(M03.Tables(0).Rows(0)("M14Criticle")) = "N" And Trim(M03.Tables(0).Rows(0)("M14grige")) = "L" Then
                                                If Trim(M01.Tables(0).Rows(Z)("M49DR_Nomal")) > 0 Then
                                                    _Operning_Capacity = (CInt(Trim(M01.Tables(0).Rows(Z)("M49DR_Nomal"))) * _CapacityHR * _LaneH * _MCNo) + _Operning_Capacity
                                                Else
                                                    _Operning_Capacity = (CInt(Trim(M01.Tables(0).Rows(Z)("M49SR_Nomal"))) * _CapacityHR * _LaneH * _MCNo) + _Operning_Capacity
                                                End If
                                            End If
                                            Y1 = Y1 + 1
                                        Next
                                        Z = Z + 1
                                    Next

                                End If
                                If dg_Dye_Detailes.Rows(2).Cells(_columcount).Text <> "" Then
                                    dg_Dye_Detailes.Rows(3).Cells(_columcount).Value = CInt(_Operning_Capacity) - dg_Dye_Detailes.Rows(2).Cells(_columcount).Value
                                Else
                                    dg_Dye_Detailes.Rows(3).Cells(_columcount).Value = CInt(_Operning_Capacity)
                                End If

                            End If
                        End If
                    End If
                    _columcount = _columcount + 1
                    userDate = userDate.AddDays(+7)

                Next

                i = i + 1
            Next
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""



        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub chkTHAN2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkTHAN2.CheckedChanged
        If chkTHAN2.Checked = True Then
            chkThan1.Checked = False
            chkLR1.Checked = False
            chkLR2.Checked = False
            chkEco1.Checked = False
            chkEco2.Checked = False
            chkE1.Checked = False
            chkE2.Checked = False
            If IsNumeric(txtDye_D_Than.Text) Then
                If txtDye_D_Than.Text > 0 Then
                    '  Call Dye_Cren2()
                    Load_Dye_Openning_Balance("THAN", txtDye_D_Than.Text)
                End If
            End If
        End If
    End Sub

    Private Sub chkThan1_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkThan1.CheckedChanged
        If chkThan1.Checked = True Then
            chkTHAN2.Checked = False
            chkLR1.Checked = False
            chkLR2.Checked = False
            chkEco1.Checked = False
            chkEco2.Checked = False
            chkE1.Checked = False
            chkE2.Checked = False
            If IsNumeric(txtDye_S_Than.Text) Then
                If txtDye_S_Than.Text > 0 Then
                    '  Call Dye_Cren2()
                    Load_Dye_Openning_Balance("THAN", txtDye_S_Than.Text)
                End If
            End If
        End If
    End Sub

    Private Sub chkEco1_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkEco1.CheckedChanged
        If chkEco1.Checked = True Then
            chkThan1.Checked = False
            chkLR1.Checked = False
            chkLR2.Checked = False
            chkTHAN2.Checked = False
            chkEco2.Checked = False
            chkE1.Checked = False
            chkE2.Checked = False
            If IsNumeric(txtDye_S_Eco.Text) Then
                If txtS_Eco1.Text > 0 Then
                    '  Call Dye_Cren2()
                    Load_Dye_Openning_Balance("ECO", txtDye_S_Eco.Text)
                End If
            End If
        End If
    End Sub

    Private Sub chkEco2_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkEco2.CheckedChanged
        If chkEco2.Checked = True Then
            chkThan1.Checked = False
            chkLR1.Checked = False
            chkLR2.Checked = False
            chkEco1.Checked = False
            chkTHAN2.Checked = False
            chkE1.Checked = False
            chkE2.Checked = False
            If IsNumeric(txtDye_D_Eco.Text) Then
                If txtDye_D_Eco1.Text > 0 Then
                    '  Call Dye_Cren2()
                    Load_Dye_Openning_Balance("ECO", txtDye_D_Eco.Text)
                End If
            End If
        End If
    End Sub

    Private Sub chkE1_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkE1.CheckedChanged
        If chkE1.Checked = True Then
            chkThan1.Checked = False
            chkLR1.Checked = False
            chkLR2.Checked = False
            chkEco1.Checked = False
            chkEco2.Checked = False
            chkTHAN2.Checked = False
            chkE2.Checked = False

            If IsNumeric(txtS_Eco1.Text) Then
                If txtS_Eco1.Text > 0 Then
                    '  Call Dye_Cren2()
                    Load_Dye_Openning_Balance("ECO +", txtS_Eco1.Text)
                End If
            End If
        End If
    End Sub

    Private Sub chkE2_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkE2.CheckedChanged
        If chkE2.Checked = True Then
            chkThan1.Checked = False
            chkLR1.Checked = False
            chkLR2.Checked = False
            chkEco1.Checked = False
            chkEco2.Checked = False
            chkTHAN2.Checked = False
            chkE1.Checked = False
            If IsNumeric(txtDye_D_Eco1.Text) Then
                If txtDye_D_Eco1.Text > 0 Then
                    '  Call Dye_Cren2()
                    Load_Dye_Openning_Balance("ECO +", txtDye_D_Eco1.Text)
                End If
            End If
        End If
    End Sub

    Private Sub BindUltraDropDown1()

        Dim dt As DataTable = New DataTable()
        ' dt.Columns.Add("ID", GetType(Integer))
        dt.Columns.Add("##", GetType(String))

        dt.Rows.Add(New Object() {"SUB"})
        dt.Rows.Add(New Object() {"APP"})
        ' dt.Rows.Add(New Object() {"SESANAL"})
        dt.AcceptChanges()

        Me.UltraDropDown4.SetDataBinding(dt, Nothing)
        '  Me.UltraDropDown1.ValueMember = "ID"
        Me.UltraDropDown4.DisplayMember = "##"
    End Sub

    Function Load_Dye_Openning_Balance(ByVal strMC As String, ByVal strQty As Integer)
        Dim vcwhere As String
        Dim i As Integer
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim M02 As DataSet
        Dim M03 As DataSet
        Dim T01 As DataSet
        Dim Value As Double
        Dim _ST As String
        Dim Y As Integer
        Dim strWeek As String
        Dim _Date As Date
        Dim _Code As Integer
        Dim Z As Integer
        Dim _WeekNo As Integer
        Dim _columcount As Integer

        Try
            i = 0
            vcwhere = "T15Sales_Order='" & strSales_Order & "' and T15Line_Item in ('" & _Dye_LineItem & "') "
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "PROX"), New SqlParameter("@vcWhereClause1", vcwhere))
            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                Dim remain As Integer
                Dim noOfWeek As Integer
                Dim userDate As Date
                Dim _LastDate As Date
                Dim _TimeSpan As TimeSpan
                Dim Z1 As Integer

                userDate = DateTime.Parse(M01.Tables(0).Rows(i)("T15Month") & "/1/" & M01.Tables(0).Rows(i)("T15Year"))
                ' MsgBox(WeekdayName(Weekday(userDate)))
                If WeekdayName(Weekday(userDate)) = "Sunday" Then
                    userDate = userDate.AddDays(-3)
                ElseIf WeekdayName(Weekday(userDate)) = "Monday" Then
                    userDate = userDate.AddDays(-4)
                ElseIf WeekdayName(Weekday(userDate)) = "Tuesday" Then
                    userDate = userDate.AddDays(-5)
                ElseIf WeekdayName(Weekday(userDate)) = "Thusday" Then
                    'userDate = userDate.AddDays(-1)
                ElseIf WeekdayName(Weekday(userDate)) = "Friday" Then
                    userDate = userDate.AddDays(-1)
                ElseIf WeekdayName(Weekday(userDate)) = "Saturday" Then
                    userDate = userDate.AddDays(-2)

                End If

                _LastDate = DateTime.Parse(M01.Tables(0).Rows(i)("T15Month") & "/1/" & M01.Tables(0).Rows(i)("T15Year"))
                ' MsgBox(Date.DaysInMonth(_LastDate.Year, _LastDate.Month))
                _LastDate = DateTime.Parse(M01.Tables(0).Rows(i)("T15Month") & "/" & Date.DaysInMonth(_LastDate.Year, _LastDate.Month) & "/" & M01.Tables(0).Rows(i)("T15Year"))
                If WeekdayName(Weekday(_LastDate)) = "Sunday" Then
                    _LastDate = _LastDate.AddDays(-4)
                ElseIf WeekdayName(Weekday(_LastDate)) = "Monday" Then
                    _LastDate = _LastDate.AddDays(-5)
                ElseIf WeekdayName(Weekday(_LastDate)) = "Tuesday" Then
                    _LastDate = _LastDate.AddDays(-6)
                ElseIf WeekdayName(Weekday(_LastDate)) = "Thusday" Then
                    'userDate = userDate.AddDays(-1)
                ElseIf WeekdayName(Weekday(_LastDate)) = "Friday" Then
                    _LastDate = _LastDate.AddDays(-2)
                ElseIf WeekdayName(Weekday(_LastDate)) = "Saturday" Then
                    _LastDate = _LastDate.AddDays(-3)

                End If


                _TimeSpan = _LastDate.Subtract(userDate)
                _WeekNo = _TimeSpan.Days / 7

                userDate = userDate.AddDays(+7)
                vcwhere = "T15Sales_Order='" & strSales_Order & "' AND t01bulk='1st Bulk' and T15Year=" & M01.Tables(0).Rows(i)("T15Year") & " and T15Month=" & M01.Tables(0).Rows(i)("T15Month") & ""
                M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "FSTB"), New SqlParameter("@vcWhereClause1", vcwhere))
                If isValidDataset(M02) Then
                    userDate = userDate.AddDays(-14)
                Else
                    userDate = userDate.AddDays(-7)
                End If
                _columcount = 1
                For Z1 = 1 To _WeekNo
                    Dim culture As System.Globalization.CultureInfo
                    Dim intWeek As Integer
                    Dim _StrWeek As String
                    Dim _StartTime As Date
                    Dim _EndTime As Date
                    Dim _CapacityHR As Integer
                    Dim _UserCapacityHR As Integer
                    Dim _OpencapacityHR As Integer
                    Dim _Baseshade As String
                    Dim _BaseMaterial As String
                    Dim characterToRemove As String
                    Dim _Operning_Capacity As Double

                    _StartTime = userDate & " " & "7:30AM"
                    _EndTime = userDate.AddDays(+7)

                    _EndTime = _EndTime & " " & "7:30AM"

                    culture = System.Globalization.CultureInfo.CurrentCulture
                    intWeek = culture.Calendar.GetWeekOfYear(userDate, System.Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Thursday)

                    dg_Dye_Detailes.Rows(2).Cells(_columcount).Value = ""
                    dg_Dye_Detailes.Rows(3).Cells(_columcount).Value = ""

                    'USED QTY 
                    'vcwhere = "tmpYear=" & Year(userDate) & " and tmpWeek=" & intWeek & " AND tmpGroup='" & strMC & "'"
                    'dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "DPB2"), New SqlParameter("@vcWhereClause1", vcwhere))
                    'If isValidDataset(dsUser) Then

                    '    dg_Dye_Detailes.Rows(2).Cells(_columcount).Value = CInt(dsUser.Tables(0).Rows(0)("Qty"))
                    'End If

                    vcwhere = "T18Year=" & Year(userDate) & " and T18WeekNo=" & intWeek & " AND T18MC='" & strMC & "'"
                    dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "DPDQ"), New SqlParameter("@vcWhereClause1", vcwhere))
                    If isValidDataset(dsUser) Then

                        dg_Dye_Detailes.Rows(2).Cells(_columcount).Value = CInt(dsUser.Tables(0).Rows(0)("Qty"))
                    End If


                    characterToRemove = "-"
                    _BaseMaterial = dgDye_Gr.Rows(0).Cells(2).Value
                    _BaseMaterial = (Replace(_BaseMaterial, characterToRemove, ""))

                    Dim _Rcode As String
                    Dim _LaneH As Integer
                    Dim Y1 As Integer
                    Dim _MCNo As Integer

                    vcwhere = "m16material='" & _BaseMaterial & "'"
                    dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "QRCD"), New SqlParameter("@vcWhereClause1", vcwhere))
                    If isValidDataset(dsUser) Then
                        _Baseshade = Trim(dsUser.Tables(0).Rows(0)("M16Shade_Type"))
                        _Rcode = Trim(dsUser.Tables(0).Rows(0)("M16R_Code"))
                    End If
                    _CapacityHR = 24 * 7 * 0.92

                    If _Baseshade = "white" Then
                        _CapacityHR = CInt(_CapacityHR / 6)
                    ElseIf _Baseshade = "Marls" Then
                        _CapacityHR = CInt(_CapacityHR / 5)
                    ElseIf _Baseshade = "Yarn Dyes" Then
                        _CapacityHR = CInt(_CapacityHR / 5)
                    Else
                        _CapacityHR = CInt(_CapacityHR / 12)
                    End If

                    _Operning_Capacity = 0
                    Z = 0
                    vcwhere = "m14order='" & _Rcode & "'" ' and m14status='CUS APPD'"
                    M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "RCOD"), New SqlParameter("@vcWhereClause1", vcwhere))
                    If isValidDataset(M02) Then
                        vcwhere = "M49quality in ('" & _DyeQuality & "') and m49mc_Group='" & strMC & "'"
                        T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "CGLI"), New SqlParameter("@vcWhereClause1", vcwhere))
                        Z = 0
                        If isValidDataset(T01) Then
                            For Each DTRow5 As DataRow In T01.Tables(0).Rows

                                Y1 = 0
                                vcwhere = "M50MC_Group='" & Trim(T01.Tables(0).Rows(Z)("M49MC_Group")) & "'"
                                dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "LAN"), New SqlParameter("@vcWhereClause1", vcwhere))
                                For Each DTRow6 As DataRow In dsUser.Tables(0).Rows
                                    _LaneH = dsUser.Tables(0).Rows(Y1)("Qty")
                                    _MCNo = dsUser.Tables(0).Rows(Y1)("MCNO")



                                    _Operning_Capacity = (strQty * _CapacityHR * _LaneH * _MCNo) + _Operning_Capacity

                                    Y1 = Y1 + 1
                                Next

                                Z = Z + 1
                            Next
                        Else '
                            _Operning_Capacity = 0
                            Y1 = 0
                            vcwhere = "M50MC_Group='" & strMC & "'"
                            dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "LAN"), New SqlParameter("@vcWhereClause1", vcwhere))
                            For Each DTRow6 As DataRow In dsUser.Tables(0).Rows
                                _LaneH = dsUser.Tables(0).Rows(Y1)("Qty")
                                _MCNo = dsUser.Tables(0).Rows(Y1)("MCNO")



                                _Operning_Capacity = (strQty * _CapacityHR * _LaneH * _MCNo) + _Operning_Capacity

                                Y1 = Y1 + 1
                            Next

                        End If
                    End If

                    ' MsgBox(Trim(dg_Dye_Detailes.Rows(2).Cells(_columcount).Text))
                    If Trim(dg_Dye_Detailes.Rows(2).Cells(_columcount).Text) <> "" Then
                        dg_Dye_Detailes.Rows(3).Cells(_columcount).Value = CInt(_Operning_Capacity) - dg_Dye_Detailes.Rows(2).Cells(_columcount).Value
                    Else
                        dg_Dye_Detailes.Rows(3).Cells(_columcount).Value = CInt(_Operning_Capacity)
                    End If
                    '=====================================================================================

                    If dg_Dye_Detailes.Rows(3).Cells(_columcount).Value > 0 Then
                        dg_Dye_Detailes.Rows(3).Cells(_columcount).Appearance.BackColor = Color.Green
                    ElseIf dg_Dye_Detailes.Rows(3).Cells(_columcount).Value = 0 Then
                        dg_Dye_Detailes.Rows(3).Cells(_columcount).Appearance.BackColor = Color.Red
                        dg_Dye_Detailes.Rows(2).Cells(_columcount).Appearance.BackColor = Color.Red
                    End If

                    If dg_Dye_Detailes.Rows(1).Cells(_columcount).Text <> "" Then
                        If dg_Dye_Detailes.Rows(1).Cells(_columcount).Value > 0 Then
                            dg_Dye_Detailes.Rows(1).Cells(_columcount).Appearance.BackColor = Color.Yellow
                        ElseIf dg_Dye_Detailes.Rows(1).Cells(_columcount).Value = 0 Then
                            dg_Dye_Detailes.Rows(1).Cells(_columcount).Appearance.BackColor = Color.Red
                            dg_Dye_Detailes.Rows(0).Cells(_columcount).Appearance.BackColor = Color.Red
                        End If
                    End If
                    _columcount = _columcount + 1
                    userDate = userDate.AddDays(+7)

                Next

                i = i + 1
            Next
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            con.close()

            '-------------------------------------------------------------


        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
                con.close()
            End If
        End Try

    End Function

    Private Sub chkLR1_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkLR1.CheckedChanged
        If chkLR1.Checked = True Then
            chkThan1.Checked = False
            chkE2.Checked = False
            chkLR2.Checked = False
            chkEco1.Checked = False
            chkEco2.Checked = False
            chkTHAN2.Checked = False
            chkE1.Checked = False
            Load_Dye_Openning_Balance("LR", txtDye_S_LR.Text)
        End If
    End Sub

    Private Sub chkLR2_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkLR2.CheckedChanged
        If chkLR2.Checked = True Then
            chkThan1.Checked = False
            chkE2.Checked = False
            chkLR1.Checked = False
            chkEco1.Checked = False
            chkEco2.Checked = False
            chkTHAN2.Checked = False
            chkE1.Checked = False
            Load_Dye_Openning_Balance("LR", txtDye_D_LR.Text)
        End If
    End Sub
    Function Dye_Shade_Gride_Main()
        On Error Resume Next
        Dim agroup1 As UltraGridGroup
        Dim agroup3 As UltraGridGroup
        Dim _MCGroup As String
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim M02 As DataSet

        Dim vcwhere As String
        Dim vcwhere1 As String
        Dim i As Integer
        Dim _LineItem As String
        Dim _Week As String
        Dim _Week1 As String
        Dim _Rowcount As Integer
        Dim _ColumCount As Integer
        Dim _WeekNo As Integer
        Dim T01 As DataSet
        Dim T02 As DataSet
        Dim tmpRow As Integer

        ' Try
        Dim Z As Integer
        dg_Dye_Shade.DisplayLayout.Bands(0).Groups.Clear()
        dg_Dye_Shade.DisplayLayout.Bands(0).Columns.Dispose()

        agroup1 = dg_Dye_Shade.DisplayLayout.Bands(0).Groups.Add("Line Item")
        agroup1.Width = 130
        Dim dt As DataTable = New DataTable()
        ' dt.Columns.Add("ID", GetType(Integer))
        Dim colWork As New DataColumn("##", GetType(String))
        dt.Columns.Add(colWork)
        colWork.ReadOnly = False


        If chkE1.Checked = True Then
            _MCGroup = "ECO"
        ElseIf chkE2.Checked = True Then
            _MCGroup = "ECO"
        ElseIf chkEco1.Checked = True Then
            _MCGroup = "ECO +"
        ElseIf chkEco2.Checked = True Then
            _MCGroup = "ECO +"
        ElseIf chkLR1.Checked = True Then
            _MCGroup = "LR"
        ElseIf chkLR2.Checked = True Then
            _MCGroup = "LR"
        ElseIf chkThan1.Checked = True Then
            _MCGroup = "THAN"
        ElseIf chkTHAN2.Checked = True Then
            _MCGroup = "THAN"
        End If


        vcwhere = "M50MC_Group='" & _MCGroup & "'"
        M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "DMLO"), New SqlParameter("@vcWhereClause1", vcwhere))
        i = 0
        _Dye_Noof_Lane = M01.Tables(0).Rows.Count
        For Each DTRow5 As DataRow In M01.Tables(0).Rows
            dt.Rows.Add(M01.Tables(0).Rows(i)("M50Lane") & " Lane")
            i = i + 1
        Next
        tmpRow = i + 1
        i = 0

        For i = 0 To 4
            dt.Rows.Add("")
        Next

        'tmpRow = i + 1
        Me.dg_Dye_Shade.SetDataBinding(dt, Nothing)
        Me.dg_Dye_Shade.DisplayLayout.Bands(0).Columns(0).Group = agroup1
        '========================================================================
        'Get Week No
        Dim tmpColum As Integer

        tmpColum = 0

        _LineItem = dg_dye_Main.Rows(0).Cells(1).Value
        vcwhere = "T10Sales_Order='" & strSales_Order & "' and T10Line_Item=" & _LineItem & " AND T10Qty>0"
        dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "GWA1"), New SqlParameter("@vcWhereClause1", vcwhere))
        i = 0
        For Each DTRow6 As DataRow In dsUser.Tables(0).Rows
            _Week = "Week " & dsUser.Tables(0).Rows(i)("M10Week")
            _Week1 = "Week- " & dsUser.Tables(0).Rows(i)("M10Week")

            agroup3 = dg_Dye_Shade.DisplayLayout.Bands(0).Groups.Add(_Week)
            agroup3.Header.Caption = _Week
            Dim Z2 As Integer
            For Z2 = 1 To 3
                If Z2 = 1 Then
                    _Week = "Overroll"
                ElseIf Z2 = 2 Then
                    _Week = "Open"
                Else
                    _Week = ""
                End If
                _Week1 = "Week- " & dsUser.Tables(0).Rows(i)("M10Week") & Z2
                Me.dg_Dye_Shade.DisplayLayout.Bands(0).Columns.Add(_Week1, _Week)
                Me.dg_Dye_Shade.DisplayLayout.Bands(0).Columns(_Week1).Group = agroup3
                Me.dg_Dye_Shade.DisplayLayout.Bands(0).Columns(_Week1).Width = 70
                Me.dg_Dye_Shade.DisplayLayout.Bands(0).Columns(_Week1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            Next

            dg_Dye_Shade.Rows(tmpRow).Cells(tmpColum).Value = "No of Batches"

            dg_Dye_Shade.Rows(tmpRow).Cells(tmpColum).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            dg_Dye_Shade.Rows(tmpRow).Cells(tmpColum).Appearance.BackColor = Color.LightGreen
            dg_Dye_Shade.Rows(tmpRow).Cells(tmpColum + 1).Appearance.BackColor = Color.Yellow
            'dg_Dye_Shade.Rows(tmpRow).Cells(tmpColum + 1).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            dg_Dye_Shade.Rows(tmpRow + 1).Cells(tmpColum).Value = "No of Full Batches"
            dg_Dye_Shade.Rows(tmpRow + 1).Cells(tmpColum).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            dg_Dye_Shade.Rows(tmpRow + 1).Cells(tmpColum).Appearance.BackColor = Color.LightGreen
            dg_Dye_Shade.Rows(tmpRow).Cells(tmpColum + 1).Appearance.BackColor = Color.Yellow
            dg_Dye_Shade.Rows(tmpRow + 1).Cells(tmpColum + 1).Appearance.BackColor = Color.Yellow
            dg_Dye_Shade.Rows(tmpRow + 2).Cells(tmpColum).Value = "Balance Qty"
            dg_Dye_Shade.Rows(tmpRow + 2).Cells(tmpColum).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            dg_Dye_Shade.Rows(tmpRow + 2).Cells(tmpColum).Appearance.BackColor = Color.LightGreen
            dg_Dye_Shade.Rows(tmpRow + 2).Cells(tmpColum + 1).Appearance.BackColor = Color.Yellow
            dg_Dye_Shade.Rows(tmpRow + 2).Cells(tmpColum + 2).Appearance.BackColor = Color.Yellow
            dg_Dye_Shade.Rows(tmpRow + 2).Cells(tmpColum + 3).Appearance.BackColor = Color.Yellow

            dg_Dye_Shade.Rows(tmpRow + 3).Cells(tmpColum).Value = "Sutable MC for Balance"
            dg_Dye_Shade.Rows(tmpRow + 3).Cells(tmpColum).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            dg_Dye_Shade.Rows(tmpRow + 3).Cells(tmpColum).Appearance.BackColor = Color.LightGreen
            dg_Dye_Shade.Rows(tmpRow + 3).Cells(tmpColum + 1).Appearance.BackColor = Color.Yellow
            dg_Dye_Shade.Rows(tmpRow + 3).Cells(tmpColum + 2).Appearance.BackColor = Color.Yellow
            dg_Dye_Shade.Rows(tmpRow + 3).Cells(tmpColum + 3).Appearance.BackColor = Color.Yellow

            tmpColum = tmpColum + 2
            dg_Dye_Shade.Rows(tmpRow).Cells(tmpColum).Value = "No of Lane"
            dg_Dye_Shade.Rows(tmpRow).Cells(tmpColum).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            dg_Dye_Shade.Rows(tmpRow).Cells(tmpColum).Appearance.BackColor = Color.LightGreen
            dg_Dye_Shade.Rows(tmpRow).Cells(tmpColum + 1).Appearance.BackColor = Color.Yellow

            dg_Dye_Shade.Rows(tmpRow + 1).Cells(tmpColum).Value = "Qty"
            dg_Dye_Shade.Rows(tmpRow + 1).Cells(tmpColum).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            dg_Dye_Shade.Rows(tmpRow + 1).Cells(tmpColum).Appearance.BackColor = Color.LightGreen
            dg_Dye_Shade.Rows(tmpRow + 1).Cells(tmpColum + 1).Appearance.BackColor = Color.Yellow
            tmpColum = tmpColum + 1
            i = i + 1
        Next
        '---------------------------------------------------------------------------------
        Me.dg_Dye_Shade.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Etched

        For i = 0 To dg_Dye_Shade.DisplayLayout.Bands(0).Columns.Count
            Me.dg_Dye_Shade.DisplayLayout.Bands(0).Columns(i).CellAppearance.BorderColor = Color.Black

        Next
        'Filing the data
        _Rowcount = 0
        _ColumCount = 1
        tmpColum = 1
        vcwhere = "M50MC_Group='" & _MCGroup & "'"
        M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "DMLO"), New SqlParameter("@vcWhereClause1", vcwhere))
        i = 0
        For Each DTRow5 As DataRow In M01.Tables(0).Rows
            Z = 0
            vcwhere = "T15Sales_Order='" & strSales_Order & "' and T15Line_Item in ('" & _Dye_LineItem & "') "
            dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "PROX"), New SqlParameter("@vcWhereClause1", vcwhere))
            For Each DTRow3 As DataRow In dsUser.Tables(0).Rows
                Dim remain As Integer
                Dim noOfWeek As Integer
                Dim userDate As Date
                Dim _LastDate As Date
                Dim _TimeSpan As TimeSpan
                Dim Z1 As Integer
                Dim _weekSt As Date

                userDate = DateTime.Parse(dsUser.Tables(0).Rows(Z)("T15Month") & "/1/" & dsUser.Tables(0).Rows(Z)("T15Year"))
                ' MsgBox(WeekdayName(Weekday(userDate)))
                If WeekdayName(Weekday(userDate)) = "Sunday" Then
                    userDate = userDate.AddDays(-3)
                ElseIf WeekdayName(Weekday(userDate)) = "Monday" Then
                    userDate = userDate.AddDays(-4)
                ElseIf WeekdayName(Weekday(userDate)) = "Tuesday" Then
                    userDate = userDate.AddDays(-5)
                ElseIf WeekdayName(Weekday(userDate)) = "Thusday" Then
                    'userDate = userDate.AddDays(-1)
                ElseIf WeekdayName(Weekday(userDate)) = "Friday" Then
                    userDate = userDate.AddDays(-1)
                ElseIf WeekdayName(Weekday(userDate)) = "Saturday" Then
                    userDate = userDate.AddDays(-2)

                End If

                _LastDate = DateTime.Parse(dsUser.Tables(0).Rows(Z)("T15Month") & "/1/" & dsUser.Tables(0).Rows(Z)("T15Year"))
                ' MsgBox(Date.DaysInMonth(_LastDate.Year, _LastDate.Month))
                _LastDate = DateTime.Parse(dsUser.Tables(0).Rows(Z)("T15Month") & "/" & Date.DaysInMonth(_LastDate.Year, _LastDate.Month) & "/" & dsUser.Tables(0).Rows(Z)("T15Year"))
                If WeekdayName(Weekday(_LastDate)) = "Sunday" Then
                    _LastDate = _LastDate.AddDays(-4)
                ElseIf WeekdayName(Weekday(_LastDate)) = "Monday" Then
                    _LastDate = _LastDate.AddDays(-5)
                ElseIf WeekdayName(Weekday(_LastDate)) = "Tuesday" Then
                    _LastDate = _LastDate.AddDays(-6)
                ElseIf WeekdayName(Weekday(_LastDate)) = "Thusday" Then
                    'userDate = userDate.AddDays(-1)
                ElseIf WeekdayName(Weekday(_LastDate)) = "Friday" Then
                    _LastDate = _LastDate.AddDays(-2)
                ElseIf WeekdayName(Weekday(_LastDate)) = "Saturday" Then
                    _LastDate = _LastDate.AddDays(-3)

                End If

                _TimeSpan = _LastDate.Subtract(userDate)
                _WeekNo = _TimeSpan.Days / 7

                userDate = userDate.AddDays(+7)
                vcwhere1 = "T15Sales_Order='" & strSales_Order & "' AND t01bulk='1st Bulk' and T15Year=" & dsUser.Tables(0).Rows(Z)("T15Year") & " and T15Month=" & dsUser.Tables(0).Rows(Z)("T15Month") & ""
                T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "FSTB"), New SqlParameter("@vcWhereClause1", vcwhere1))
                If isValidDataset(T01) Then
                    userDate = userDate.AddDays(-14)
                Else
                    userDate = userDate.AddDays(-7)
                End If
                _ColumCount = 1

                For Z1 = 1 To _WeekNo
                    Dim culture As System.Globalization.CultureInfo
                    Dim intWeek As Integer
                    Dim _StrWeek As String
                    Dim _StartTime As Date
                    Dim _EndTime As Date
                    Dim _CapacityHR As Integer
                    Dim _UserCapacityHR As Integer
                    Dim _OpencapacityHR As Integer
                    Dim _Baseshade As String
                    Dim _BaseMaterial As String
                    Dim characterToRemove As String
                    Dim _Operning_Capacity As Double

                    _StartTime = userDate & " " & "7:30AM"
                    _EndTime = userDate.AddDays(+7)

                    _EndTime = _EndTime & " " & "7:30AM"

                    'If Z = 1 Then
                    '    _weekSt = userDate
                    'Else
                    ' _weekSt = userDate.AddDays(+6)
                    ' End If
                    culture = System.Globalization.CultureInfo.CurrentCulture
                    intWeek = culture.Calendar.GetWeekOfYear(userDate, System.Globalization.CalendarWeekRule.FirstFullWeek, DayOfWeek.Thursday)

                    '--------------------------------------
                    vcwhere = "T10Sales_Order='" & strSales_Order & "' and T10Line_Item=" & _LineItem & " AND T10Qty>0 and M10Week=" & intWeek & ""
                    T02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "GWA1"), New SqlParameter("@vcWhereClause1", vcwhere))
                    If isValidDataset(T02) Then

                        ' _Week = "Week " & dsUser.Tables(0).Rows(i)("M10Week")
                        'Over roll capacity
                        vcwhere = "tmpGroup='" & _MCGroup & "' and tmpweek=" & intWeek & " and tmpyear='" & Year(_StartTime) & "' and tmpCapacity=" & M01.Tables(0).Rows(i)("M50Lane") & ""
                        T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "DPB2"), New SqlParameter("@vcWhereClause1", vcwhere))
                        If isValidDataset(T01) Then
                            dg_Dye_Shade.Rows(_Rowcount).Cells(_ColumCount).Value = CInt(T01.Tables(0).Rows(0)("Qty"))
                        End If
                        _ColumCount = _ColumCount + 1
                        '---------------------------------------------------------
                        'Total Capacity
                        characterToRemove = "-"
                        _BaseMaterial = dgDye_Gr.Rows(0).Cells(2).Value
                        _BaseMaterial = (Replace(_BaseMaterial, characterToRemove, ""))

                        vcwhere = "m16Quality='" & _DyeQuality & "' and m16material='" & _BaseMaterial & "'"
                        dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "QRCD"), New SqlParameter("@vcWhereClause1", vcwhere))
                        If isValidDataset(dsUser) Then
                            _Baseshade = Trim(dsUser.Tables(0).Rows(0)("M16Shade_Type"))
                        End If
                        _CapacityHR = 24 * 7 * 0.92

                        If _Baseshade = "white" Then
                            _CapacityHR = CInt(_CapacityHR / 6)
                        ElseIf _Baseshade = "Marls" Then
                            _CapacityHR = CInt(_CapacityHR / 5)
                        ElseIf _Baseshade = "Yarn Dyes" Then
                            _CapacityHR = CInt(_CapacityHR / 5)
                        Else
                            _CapacityHR = CInt(_CapacityHR / 12)
                        End If

                        _Operning_Capacity = 0

                        vcwhere = "M14quality='" & _DyeQuality & "'" ' and m14status='CUS APPD'"
                        M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "RCOD"), New SqlParameter("@vcWhereClause1", vcwhere))
                        If isValidDataset(M02) Then
                            'vcwhere = "M49quality='" & _DyeQuality & "' and m49mc_Group='" & _MCGroup & "'"
                            'T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "CGLI"), New SqlParameter("@vcWhereClause1", vcwhere))
                            'Z = 0
                            'For Each DTRow6 As DataRow In T01.Tables(0).Rows
                            Dim _LaneH As Integer

                            '    'vcwhere = "M50MC_Group='" & Trim(T01.Tables(0).Rows(Z)("M49MC_Group")) & "'"
                            '    'dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "LAN"), New SqlParameter("@vcWhereClause1", vcwhere))
                            'If isValidDataset(dsUser) Then
                            '    _LaneH = dsUser.Tables(0).Rows(0)("Qty")
                            'End If

                            _LaneH = M01.Tables(0).Rows(i)("M50Lane")

                            strQty = 0
                            '        For Z = 1 To dg_Dye_Detailes.DisplayLayout.Bands(0).Columns.Count - 1
                            '            MsgBox(dg_Dye_Detailes.Rows(4).Cells(Z).Text)
                            '            strQty = 0
                            '            If Microsoft.VisualBasic.Right(dg_Dye_Detailes.DisplayLayout.Bands(0).Columns(Z).Header.Caption, 2) = intWeek Then
                            '                strQty = dg_Dye_Detailes.Rows(4).Cells(Z).Text

                            '                Exit For

                            '            End If
                            'Next

                            If chkThan1.Checked = True Then
                                strQty = txtDye_S_Than.Text
                            ElseIf chkTHAN2.Checked = True Then
                                strQty = txtDye_D_Than.Text
                            ElseIf chkLR1.Checked = True Then
                                strQty = txtDye_S_LR.Text
                            ElseIf chkLR2.Checked = True Then
                                strQty = txtDye_D_LR.Text
                            ElseIf chkEco1.Checked = True Then
                                strQty = txtDye_S_Eco.Text

                            ElseIf chkEco2.Checked = True Then
                                strQty = txtDye_D_Eco1.Text
                            ElseIf chkE1.Checked = True Then
                                strQty = txtS_Eco1.Text
                            ElseIf chkE2.Checked = True Then
                                strQty = txtDye_D_Eco.Text
                            End If
                            _Operning_Capacity = (strQty * _CapacityHR * _LaneH * M01.Tables(0).Rows(i)("MCNo")) + _Operning_Capacity

                            If Microsoft.VisualBasic.Right(dg_Dye_Shade.DisplayLayout.Bands(0).Columns(Z).Header.Caption, 2) = intWeek Then
                                If Trim(dg_Dye_Shade.Rows(_Rowcount).Cells(_ColumCount - 1).Text) <> "" Then
                                    dg_Dye_Shade.Rows(_Rowcount).Cells(_ColumCount).Value = CInt(_Operning_Capacity) - dg_Dye_Shade.Rows(_Rowcount).Cells(_ColumCount - 1).Value
                                Else
                                    dg_Dye_Shade.Rows(_Rowcount).Cells(_ColumCount).Value = CInt(_Operning_Capacity)
                                End If
                                _ColumCount = _ColumCount + 1
                                dg_Dye_Shade.Rows(_Rowcount).Cells(_ColumCount).Style = ColumnStyle.CheckBox
                                dg_Dye_Shade.Rows(_Rowcount).Cells(_ColumCount).Value = False
                                _ColumCount = _ColumCount + 1
                                'Z = Z + 1
                                'Next
                            End If

                        End If
                    End If

                    vcwhere = "T10Sales_Order='" & strSales_Order & "' and T10Line_Item=" & _LineItem & " and M10Week=" & intWeek & " and T10Qty>0"
                    M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "SDMA"), New SqlParameter("@vcWhereClause1", vcwhere))
                    If isValidDataset(M02) Then
                        dg_Dye_Shade.Rows(tmpRow + 1).Cells(tmpColum + 2).Value = CInt(M02.Tables(0).Rows(0)("T10Qty"))
                        dg_Dye_Shade.Rows(tmpRow).Cells(tmpColum + 2).Value = CInt(M02.Tables(0).Rows(0)("T10Qty")) / strQty

                        tmpColum = tmpColum + 3
                    End If

                    userDate = userDate.AddDays(+7)

                Next

                Z = Z + 1
            Next
            _Rowcount = _Rowcount + 1
            i = i + 1
        Next

        '===========================================

        DBEngin.CloseConnection(con)
        con.ConnectionString = ""

        'Catch ex As EvaluateException
        '    If transactionCreated = False Then transaction.Rollback()
        '    MessageBox.Show(Me, ex.ToString)

        'Finally
        '    If connectionCreated Then DBEngin.CloseConnection(connection)
        'End Try


    End Function

    Function Insert_T10Dye_Week_Allocation_MCGroup(ByVal strCount As Integer, ByVal strWeek As Integer, ByVal strQty As Double, ByVal strSub As String, ByVal strApp As String)
        Dim vcWhere As String
        Dim M01 As DataSet
        Dim vcFieldList As String
        Dim ncQryType As String
        Dim nvcFieldList1 As String
        Dim _MCGroup As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Dim _LineItem As String
        Dim _UpdateStatus As Boolean

        Try
            If chkE1.Checked = True Then
                _MCGroup = "ECO"
            ElseIf chkE2.Checked = True Then
                _MCGroup = "ECO"
            ElseIf chkEco1.Checked = True Then
                _MCGroup = "ECO +"
            ElseIf chkEco2.Checked = True Then
                _MCGroup = "ECO +"
            ElseIf chkLR1.Checked = True Then
                _MCGroup = "LR"
            ElseIf chkLR2.Checked = True Then
                _MCGroup = "LR"
            ElseIf chkThan1.Checked = True Then
                _MCGroup = "THAN"
            ElseIf chkTHAN2.Checked = True Then
                _MCGroup = "THAN"
            End If
            _LineItem = dg_dye_Main.Rows(0).Cells(1).Value
            vcWhere = "T10Sales_Order='" & strSales_Order & "' and T10Line_Item=" & _LineItem & " and M10Week=" & strWeek & ""
            M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "SDMA"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                vcFieldList = "update T10Dye_Week_Allocation_MCGroup set T10MC_Group='" & _MCGroup & "',T10Qty='" & strQty & "',T10Submision='" & strSub & "',T10Approve='" & strApp & "' where T10Sales_Order='" & strSales_Order & "' and T10Line_Item=" & _LineItem & " and M10Week=" & strWeek & ""
                ExecuteNonQueryText(connection, transaction, vcFieldList)
                ' Call Dye_Shade_Gride_Main(
                _UpdateStatus = True
            Else
                If strQty > 0 Then
                    ncQryType = "DMAG"
                    nvcFieldList1 = "(T18Ref_No," & "T18O_No," & "T18Date," & "T10Sales_Order," & "T10Line_Item," & "T10MC_Group," & "M10Week," & "T10Qty," & "T10User," & "T10Submision," & "T10Approve) " & "values(" & Delivary_Ref & "," & strCount & ",'" & Today & "','" & strSales_Order & "'," & _LineItem & ",'" & _MCGroup & "'," & strWeek & ",'" & strQty & "','" & strDisname & "','" & strSub & "','" & strApp & "')"
                    up_GetSetYarn_Bookingtmp(ncQryType, nvcFieldList1, vcWhere, connection, transaction)
                    _UpdateStatus = True
                End If
                End If
            transaction.Commit()
            
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
            connection.Close()
            If _UpdateStatus = True Then
                Call Dye_Shade_Gride_Main()
            End If

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)

                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
                connection.Close()
            End If
        End Try

    End Function

    Private Sub dg_Dye_Detailes_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles dg_Dye_Detailes.AfterCellUpdate
        Call Calculate_Dye_Balance()

    End Sub


    Private Sub cmdDye_Left_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDye_Left.Click
        Panel17.Visible = False
        With UltraGroupBox35
            .Location = New Point(410, 225)
            .Width = 719
            .Height = 276
        End With

        dg_Dye_Shade.Width = 680
        dg_Dye_Shade.Height = 270
        UltraLabel103.Visible = False
        lblPrevious_St_Code.Visible = False
        UltraLabel81.Visible = False
        lblConstrain.Visible = False
    End Sub


    Private Sub dg_Dye_Detailes_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles dg_Dye_Detailes.InitializeLayout

    End Sub

    Private Sub UltraButton25_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton25.Click
        On Error Resume Next
        Dim i As Integer
        Dim _Qty As Double
        Dim _CellIndex As Integer
        Dim _WeekNo As Integer
        Dim _WeekQty As Double
        Dim _Sub As String
        Dim _App As String
        i = 0
        ' MsgBox(dg_Dye_Detailes.DisplayLayout.Bands(0).Columns.Count)
        For i = 1 To dg_Dye_Detailes.DisplayLayout.Bands(0).Columns.Count - 1
            _Sub = "N"
            _App = "N"
            _WeekQty = 0
            If IsNumeric(dg_Dye_Detailes.Rows(4).Cells(i).Value) Then
                If dg_Dye_Detailes.Rows(4).Cells(i).Value > CDbl(lblDye_Qty.Text) Then
                    Dim windowInfo As New Infragistics.Win.Misc.UltraDesktopAlertShowWindowInfo
                    Dim strFileName As String
                    windowInfo.Caption = "Order Quantity less than Allocated qty."
                    windowInfo.FooterText = "Technova"
                    strFileName = ConfigurationManager.AppSettings("SoundPath") + "\REMINDER.wav"
                    windowInfo.Sound = strFileName
                    UltraDesktopAlert1.Show(windowInfo)
                    Exit Sub
                Else
                    'INSERT TABLE
                    _WeekNo = Microsoft.VisualBasic.Right(dg_Dye_Detailes.DisplayLayout.Bands(0).Columns(i).Header.Caption, 2)
                    If dg_Dye_Detailes.Rows(4).Cells(i).Text = "SUB" Then
                        _WeekQty = 0
                        _Sub = "Y"
                    ElseIf dg_Dye_Detailes.Rows(4).Cells(i).Text = "APP" Then
                        _WeekQty = 0
                        _App = "Y"
                    End If

                    _WeekQty = dg_Dye_Detailes.Rows(4).Cells(i).Value
                    Call Insert_T10Dye_Week_Allocation_MCGroup(i, _WeekNo, _WeekQty, _Sub, _App)

                End If
            Else
                _WeekNo = Microsoft.VisualBasic.Right(dg_Dye_Detailes.DisplayLayout.Bands(0).Columns(i).Header.Caption, 2)
                _WeekQty = 0
                _Sub = "N"
                _Sub = "N"
                If dg_Dye_Detailes.Rows(4).Cells(i).Text = "SUB" Then
                    _WeekQty = 0
                    _Sub = "Y"
                ElseIf dg_Dye_Detailes.Rows(4).Cells(i).Text = "APP" Then
                    _WeekQty = 0
                    _App = "Y"
                End If
                Call Insert_T10Dye_Week_Allocation_MCGroup(i, _WeekNo, _WeekQty, _Sub, _App)
            End If
            If dg_Dye_Detailes.Rows(5).Cells(i).Value = True Then
                If IsNumeric(dg_Dye_Detailes.Rows(3).Cells(i).Value) Then
                    If dg_Dye_Detailes.Rows(3).Cells(i).Value > CDbl(lblDye_Qty.Text) Then
                        Dim windowInfo As New Infragistics.Win.Misc.UltraDesktopAlertShowWindowInfo
                        Dim strFileName As String
                        windowInfo.Caption = "Order Quantity less than Allocated qty."
                        windowInfo.FooterText = "Technova"
                        strFileName = ConfigurationManager.AppSettings("SoundPath") + "\REMINDER.wav"
                        windowInfo.Sound = strFileName
                        UltraDesktopAlert1.Show(windowInfo)
                        Exit Sub
                    Else
                        dg_Dye_Detailes.Rows(4).Cells(i).Value = dg_Dye_Detailes.Rows(3).Cells(i).Value
                        'INSER TABLE
                        _WeekQty = dg_Dye_Detailes.Rows(3).Cells(i).Value
                        _WeekNo = Microsoft.VisualBasic.Right(dg_Dye_Detailes.DisplayLayout.Bands(0).Columns(i).Header.Caption, 2)
                        If dg_Dye_Detailes.Rows(3).Cells(i).Text = "SUB" Then
                            _WeekQty = 0
                            _Sub = "Y"
                        ElseIf dg_Dye_Detailes.Rows(3).Cells(i).Text = "APP" Then
                            _WeekQty = 0
                            _App = "Y"
                        End If

                        Call Insert_T10Dye_Week_Allocation_MCGroup(i, _WeekNo, _WeekQty, _Sub, _App)
                    End If
                End If
            End If
        Next
        ' MsgBox(dg_Dye_Detailes.Rows(4).Cells.GetItem)
        '_CellIndex = dg_Dye_Detailes.ActiveCell.Column.Key
        'If IsNumeric(dg_Dye_Detailes.Rows(4).Cells(_CellIndex).Value) Then
        '    If CDbl(dg_Dye_Detailes.Rows(4).Cells(_CellIndex).Value) > CDbl(lblDye_Qty.Text) Then
        '        Dim windowInfo As New Infragistics.Win.Misc.UltraDesktopAlertShowWindowInfo
        '        Dim strFileName As String
        '        windowInfo.Caption = "Order Quantity less than Allocated qty."
        '        windowInfo.FooterText = "Technova"
        '        strFileName = ConfigurationManager.AppSettings("SoundPath") + "\REMINDER.wav"
        '        windowInfo.Sound = strFileName
        '        UltraDesktopAlert1.Show(windowInfo)
        '        Exit Sub
        '    End If
        'End If

        'For i = 1 To dg_Dye_Detailes.DisplayLayout.Bands(0).Columns.Count
        '    If IsNumeric(dg_Dye_Detailes.Rows(5).Cells(i).Value) Then
        '        _Qty = _Qty + dg_Dye_Detailes.Rows(5).Cells(i).Value
        '    End If
        'Next
    End Sub

    Private Sub UltraButton26_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton26.Click
        On Error Resume Next

        Dim _RowCount As Integer
        Dim _ColoumCount As Integer
        Dim tmpRow As Integer
        Dim tmpColum As Integer
        Dim i As Integer
        Dim Z As Integer
        Dim _Barch As Integer
        Dim _lane As Integer
        Dim _Capacity As Integer
        Dim _MCGroup As String
        _RowCount = 0
        _ColoumCount = 1

        If CDbl(lblDye_Balance.Text) > 0 Then
            MsgBox("Please allocate complete Qty", MsgBoxStyle.Information, "Information ......")
            Exit Sub
        End If
        If chkThan1.Checked = True Then
            _Capacity = txtDye_S_Than.Text
            _MCGroup = "THAN"
        ElseIf chkTHAN2.Checked = True Then
            _Capacity = txtDye_D_Than.Text
            _MCGroup = "THAN"
        ElseIf chkLR1.Checked = True Then
            _Capacity = txtDye_S_LR.Text
            _MCGroup = "LR"
        ElseIf chkLR2.Checked = True Then
            _Capacity = txtDye_D_LR.Text
            _MCGroup = "LR"
        ElseIf chkE1.Checked = True Then
            _Capacity = txtS_Eco1.Text
            _MCGroup = "ECO"
        ElseIf chkE2.Checked = True Then
            _Capacity = txtDye_D_Eco1.Text
            _MCGroup = "ECO"
        ElseIf chkEco1.Checked = True Then
            _Capacity = txtDye_S_Eco.Text
            _MCGroup = "ECO +"
        ElseIf chkEco2.Checked = True Then
            _Capacity = txtDye_D_Eco.Text
            _MCGroup = "ECO +"
        End If
        If Save_Block_DyeMC(_MCGroup) = True Then
            Exit Sub
        End If

        Dye_Cal_Status = False
        tmpRow = _Dye_Noof_Lane + 1
        For i = 0 To _Dye_Noof_Lane
            Z = 0
            _ColoumCount = 1
            ' MsgBox(dg_Dye_Shade.DisplayLayout.Bands(0).Columns.Count)
            For Z = 1 To dg_Dye_Shade.DisplayLayout.Bands(0).Columns.Count - 1
                ' MsgBox(Microsoft.VisualBasic.Left(dg_Dye_Shade.Rows(_RowCount).Cells(_ColoumCount - 1).Text, 1))
                If Trim(dg_Dye_Shade.Rows(_RowCount).Cells(_ColoumCount).Text) = "True" Then
                    _Barch = (Microsoft.VisualBasic.Left(dg_Dye_Shade.Rows(_RowCount).Cells(0).Value, 1))
                    'MsgBox(dg_Dye_Shade.Rows(tmpRow).Cells(_ColoumCount).Value / _Barch)
                    dg_Dye_Shade.Rows(tmpRow).Cells(_ColoumCount - 2).Value = (dg_Dye_Shade.Rows(tmpRow).Cells(_ColoumCount).Value) / _Barch
                    dg_Dye_Shade.Rows(tmpRow + 1).Cells(_ColoumCount - 2).Value = Fix((dg_Dye_Shade.Rows(tmpRow).Cells(_ColoumCount).Value) / _Barch)
                    If Fix(dg_Dye_Shade.Rows(tmpRow + 1).Cells(_ColoumCount - 2).Value) = 0 Then
                        '   MsgBox(dg_Dye_Shade.Rows(tmpRow + 1).Cells(_ColoumCount).Value)
                        dg_Dye_Shade.Rows(tmpRow + 1).Cells(_ColoumCount - 2).Value = "1"
                        dg_Dye_Shade.Rows(tmpRow + 2).Cells(_ColoumCount - 2).Value = "0" 'dg_Dye_Shade.Rows(tmpRow + 1).Cells(_ColoumCount).Value - (1 * _Capacity)
                    Else

                        dg_Dye_Shade.Rows(tmpRow + 2).Cells(_ColoumCount - 2).Value = dg_Dye_Shade.Rows(tmpRow + 1).Cells(_ColoumCount).Value - (dg_Dye_Shade.Rows(tmpRow + 1).Cells(_ColoumCount - 2).Value * _Capacity)
                    End If
                Else

                End If
                _ColoumCount = _ColoumCount + 1
            Next
            _RowCount = _RowCount + 1
            Dye_Cal_Status = True
        Next

        'For i = 0 To dg_Dye_Shade.DisplayLayout.Bands(0).Columns.Count
        '    'MsgBox(Microsoft.VisualBasic.Left(dg_Dye_Shade.Rows(_RowCount).Cells(0).Text, 1))
        '    If IsNumeric(Microsoft.VisualBasic.Left(dg_Dye_Shade.Rows(_RowCount).Cells(0).Text, 1)) Then
        '        _lane = Microsoft.VisualBasic.Left(dg_Dye_Shade.Rows(_RowCount).Cells(0).Text, 1)
        '        If dg_Dye_Shade.Rows(_RowCount).Cells(_ColoumCount + 2).Text = True Then
        '            _ColoumCount = _ColoumCount + 2

        '            Continue For
        '        End If


        '    End If
        'Next
        Call Finding_Sutable_MCGroup(_MCGroup, _Capacity)
        cmdD_Save.Enabled = True

    End Sub

    Function Calculate_Dye_Balance()
        On Error Resume Next
        Dim Value As Double
        Dim _ColCount As Integer
        Dim _Qty As Double
        _ColCount = 1
        _Qty = 0
        For _ColCount = 1 To dg_Dye_Detailes.DisplayLayout.Bands(0).Columns.Count - 1
            If IsNumeric(dg_Dye_Detailes.Rows(4).Cells(_ColCount).Value) Then
                _Qty = _Qty + dg_Dye_Detailes.Rows(4).Cells(_ColCount).Value

            End If
            '_ColCount = _ColCount + 1
        Next
        Value = CDbl(lblDye_Qty.Text) - _Qty
        lblDye_Balance.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
        lblDye_Balance.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
    End Function

    Function Finding_Sutable_MCGroup(ByVal strMC_Group As String, ByVal strCapacity As Double)
        Dim i As Integer
        Dim vcWhere As String
        Dim M01 As DataSet
        Dim M02 As DataSet
        Dim Value As Double
        Dim _VString As String
        Dim Diff As TimeSpan
        Dim _To As Date
        Dim MyText As String
        Dim _Dg1_Colume As Integer
        Dim _dg2_coloume As Integer
        Dim _weekNo As Integer
        Dim Z As Integer
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim _WeekDis As String
        Dim vcwhere1 As String
        Dim T01 As DataSet
        Dim _LastRow As Integer

        _Dg1_Colume = 1
        ' For _Dg1_Colume=1 to dg_Dye_Detailes.
        _Dg1_Colume = 1
        'For _Dg1_Colume = 1 To dg_Dye_Detailes.DisplayLayout.Bands(0).Columns.Count - 1
        '    '_weekNo = dg_Dye_Detailes.DisplayLayout.Bands(0).Columns(_Dg1_Colume)
        '    ' MsgBox(dg_Dye_Detailes.Rows(4).Cells(_Dg1_Colume).Value)
        '    ' MsgBox(dg_Dye_Detailes.Layouts(0).Bands(0).Columns(0).Header.Caption)

        '    If IsNumeric(dg_Dye_Detailes.Rows(4).Cells(_Dg1_Colume).Value) Then
        '        _WeekDis = dg_Dye_Detailes.DisplayLayout.Bands(0).Columns(_Dg1_Colume).Header.Caption
        '        _weekNo = Microsoft.VisualBasic.Right(_WeekDis, 3)
        '    End If
        'Next
        _LastRow = dg_Dye_Shade.Rows.Count
        Z = 0
        vcWhere = "T15Sales_Order='" & strSales_Order & "' and T15Line_Item in ('" & _Dye_LineItem & "') "
        dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "PROX"), New SqlParameter("@vcWhereClause1", vcWhere))
        For Each DTRow3 As DataRow In dsUser.Tables(0).Rows
            Dim remain As Integer
            Dim noOfWeek As Integer
            Dim userDate As Date
            Dim _LastDate As Date
            Dim _TimeSpan As TimeSpan
            Dim Z1 As Integer
            Dim _weekSt As Date

            userDate = DateTime.Parse(dsUser.Tables(0).Rows(Z)("T15Month") & "/1/" & dsUser.Tables(0).Rows(Z)("T15Year"))
            ' MsgBox(WeekdayName(Weekday(userDate)))
            If WeekdayName(Weekday(userDate)) = "Sunday" Then
                userDate = userDate.AddDays(-3)
            ElseIf WeekdayName(Weekday(userDate)) = "Monday" Then
                userDate = userDate.AddDays(-4)
            ElseIf WeekdayName(Weekday(userDate)) = "Tuesday" Then
                userDate = userDate.AddDays(-5)
            ElseIf WeekdayName(Weekday(userDate)) = "Thusday" Then
                'userDate = userDate.AddDays(-1)
            ElseIf WeekdayName(Weekday(userDate)) = "Friday" Then
                userDate = userDate.AddDays(-1)
            ElseIf WeekdayName(Weekday(userDate)) = "Saturday" Then
                userDate = userDate.AddDays(-2)

            End If

            _LastDate = DateTime.Parse(dsUser.Tables(0).Rows(Z)("T15Month") & "/1/" & dsUser.Tables(0).Rows(Z)("T15Year"))
            ' MsgBox(Date.DaysInMonth(_LastDate.Year, _LastDate.Month))
            _LastDate = DateTime.Parse(dsUser.Tables(0).Rows(Z)("T15Month") & "/" & Date.DaysInMonth(_LastDate.Year, _LastDate.Month) & "/" & dsUser.Tables(0).Rows(Z)("T15Year"))
            If WeekdayName(Weekday(_LastDate)) = "Sunday" Then
                _LastDate = _LastDate.AddDays(-4)
            ElseIf WeekdayName(Weekday(_LastDate)) = "Monday" Then
                _LastDate = _LastDate.AddDays(-5)
            ElseIf WeekdayName(Weekday(_LastDate)) = "Tuesday" Then
                _LastDate = _LastDate.AddDays(-6)
            ElseIf WeekdayName(Weekday(_LastDate)) = "Thusday" Then
                'userDate = userDate.AddDays(-1)
            ElseIf WeekdayName(Weekday(_LastDate)) = "Friday" Then
                _LastDate = _LastDate.AddDays(-2)
            ElseIf WeekdayName(Weekday(_LastDate)) = "Saturday" Then
                _LastDate = _LastDate.AddDays(-3)

            End If

            _TimeSpan = _LastDate.Subtract(userDate)
            _weekNo = _TimeSpan.Days / 7

            userDate = userDate.AddDays(+7)
            vcwhere1 = "T15Sales_Order='" & strSales_Order & "' AND t01bulk='1st Bulk' and T15Year=" & dsUser.Tables(0).Rows(Z)("T15Year") & " and T15Month=" & dsUser.Tables(0).Rows(Z)("T15Month") & ""
            T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "FSTB"), New SqlParameter("@vcWhereClause1", vcwhere1))
            If isValidDataset(T01) Then
                userDate = userDate.AddDays(-14)
            Else
                userDate = userDate.AddDays(-7)
            End If



            _Dg1_Colume = 1
            ' For _Dg1_Colume=1 to dg_Dye_Detailes.
            _Dg1_Colume = 1

            For Z1 = 1 To _weekNo
                Dim culture As System.Globalization.CultureInfo
                Dim intWeek As Integer
                Dim _StrWeek As String
                Dim _StartTime As Date
                Dim _EndTime As Date
                Dim _CapacityHR As Integer
                Dim _UserCapacityHR As Integer
                Dim _OpencapacityHR As Integer
                Dim _Baseshade As String
                Dim _BaseMaterial As String
                Dim characterToRemove As String
                Dim _Operning_Capacity As Double

                _StartTime = userDate & " " & "7:30AM"
                _EndTime = userDate.AddDays(+7)

                _EndTime = _EndTime & " " & "7:30AM"

                'If Z = 1 Then
                '    _weekSt = userDate
                'Else
                ' _weekSt = userDate.AddDays(+6)
                ' End If
                culture = System.Globalization.CultureInfo.CurrentCulture
                intWeek = culture.Calendar.GetWeekOfYear(userDate, System.Globalization.CalendarWeekRule.FirstFullWeek, DayOfWeek.Thursday)

                'MsgBox(dg_Dye_Shade.Rows(_LastRow - 2).Cells(_Dg1_Colume).Text)
                If IsNumeric(dg_Dye_Shade.Rows(_LastRow - 2).Cells(_Dg1_Colume).Text) Then
                    vcWhere = "m50Mc_Group='" & strMC_Group & "' and M50Min_Loading<='" & dg_Dye_Shade.Rows(_LastRow - 2).Cells(_Dg1_Colume).Value & "'"
                    dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "DMCC"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(dsUser) Then
                        dg_Dye_Shade.Rows(_LastRow - 1).Cells(_Dg1_Colume).Value = dsUser.Tables(0).Rows(0)("M50Mc_No")
                    Else
                        If dg_Dye_Shade.Rows(_LastRow - 2).Cells(_Dg1_Colume).Text > 0 Then
                            vcWhere = "m50Mc_Group='" & strMC_Group & "' "
                            dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "DMCC"), New SqlParameter("@vcWhereClause1", vcWhere))
                            If isValidDataset(dsUser) Then
                                dg_Dye_Shade.Rows(_LastRow - 1).Cells(_Dg1_Colume).Value = dsUser.Tables(0).Rows(0)("M50Mc_No")
                            End If
                        End If
                        End If
                End If

                _Dg1_Colume = _Dg1_Colume + 3
                userDate = userDate.AddDays(+7)
            Next
            Z = Z + 1
        Next
    End Function

    Function Save_Block_DyeMC(ByVal strMC_Group As String) As Boolean
        Dim vcWhere As String
        Dim M01 As DataSet
        Dim vcFieldList As String
        Dim ncQryType As String
        Dim nvcFieldList1 As String
        Dim M02 As DataSet

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean
        Try
            connection = DBEngin.GetConnection(True)
            connectionCreated = True
            transaction = connection.BeginTransaction()
            transactionCreated = True

            nvcFieldList1 = "delete from tmpBlock_Dye_Machine where tmpUser='" & strDisname & "'"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            vcFieldList = "tmpMC_Group='" & Trim(strMC_Group) & "'"
            M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "DMC1"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                Save_Block_DyeMC = True
                MsgBox("This Dye Machine Grop used by " & Trim(M01.Tables(0).Rows(0)("tmpUser")), MsgBoxStyle.Information, "Information ......")
                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""

                Exit Function
            Else
                ncQryType = "TBDM"
                nvcFieldList1 = "(tmpRef_No," & "tmpSales_Order," & "tmpLine_Item," & "tmpMC_Group," & "tmpDate," & "tmpUser) " & "values(" & Delivary_Ref & ",'" & strSales_Order & "','" & strLine_Item & "','" & strMC_Group & "','" & Now & "','" & strDisname & "')"
                up_GetSetCAPACITY(ncQryType, nvcFieldList1, vcWhere, connection, transaction)

            End If
            transaction.Commit()
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""

        Catch ex As EvaluateException
            If transactionCreated = False Then transaction.Rollback()
            MessageBox.Show(Me, ex.ToString)

        Finally
            If connectionCreated Then DBEngin.CloseConnection(connection)
        End Try

    End Function

    Private Sub cmdDye_Right_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdDye_Right.Click
        Panel17.Visible = True
        With UltraGroupBox35
            .Location = New Point(710, 225)
            .Width = 352
            .Height = 216
        End With
        dg_Dye_Shade.Width = 306
        dg_Dye_Shade.Height = 186

        UltraLabel103.Visible = True
        lblPrevious_St_Code.Visible = True
        UltraLabel81.Visible = True
        lblConstrain.Visible = True
    End Sub

    Private Sub cmdYarn_Booking_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdYarn_Booking.Click
        ' On Error Resume Next
        Dim i As Integer
        Dim _Status As Boolean
        i = 0
        Call Load_Dye_YarnCombo()
        If UltraGroupBox37.Visible = True Then
            UltraGroupBox37.Visible = False
        Else
            _Status = True
            For Each uRow As UltraGridRow In dg1.Rows
                If dg1.Rows(i).Cells(0).Text <> "" Then
                    If dg1.Rows(i).Cells(9).Text <> "" Then
                    Else
                        _Status = False
                        Exit For
                    End If
                End If
                i = i + 1
            Next

            If _Status = True Then
                MsgBox("No need to Order the yarn", MsgBoxStyle.Information, "Information ......")
            Else
                UltraGroupBox37.Visible = True
                Call Load_Gride_Dyeyarn_Request()
            End If
        End If
    End Sub

    Function Load_Gride_Dyeyarn_Request()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer_DyeYarn = CustomerDataClass.MakeDataTable_Dye_YarnRequest
        dgDYYarn_Request.DataSource = c_dataCustomer_DyeYarn
        With dgDYYarn_Request
            .DisplayLayout.Bands(0).Columns(0).Width = 170
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 50
            .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(3).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 60
            .DisplayLayout.Bands(0).Columns(3).Width = 70
            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
        End With

    End Function


    Function Load_Gride_yarn_Request()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer_Yarn = CustomerDataClass.MakeDataTable_Dye_YarnRequest
        dg_Yarn_Request.DataSource = c_dataCustomer_Yarn
        With dg_Yarn_Request
            .DisplayLayout.Bands(0).Columns(0).Width = 170
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 50
            .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(3).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 60
            .DisplayLayout.Bands(0).Columns(3).Width = 70
            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
        End With

    End Function

    Private Sub UltraButton27_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton27.Click
        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True


        Dim vcWhere As String
        Dim M01 As DataSet
        Dim ncQryType As String
        Dim i As Integer
        Dim _strUserName As String
        Dim _strEmail As String

        i = 0

        Try
            Cursor = Cursors.WaitCursor
            For Each uRow As UltraGridRow In UltraGrid1.Rows
                ' MsgBox(UltraGrid1.Rows(i).Cells(5).Value)
                If IsNumeric(UltraGrid1.Rows(i).Cells(5).Value) Then
                    vcWhere = "T12Ref_No=" & Delivary_Ref & " and T12Sales_Order='" & strSales_Order & "' and T12Line_Item=" & strLine_Item & " and T12Stock_Code='" & Trim(UltraGrid1.Rows(i).Cells(2).Value) & "'"
                    M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetUse_Griege_Qty", New SqlParameter("@cQryType", "CGS1"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M01) Then
                    Else
                        ncQryType = "GADD"
                        nvcFieldList1 = "(T12Ref_No," & "T12Sales_Order," & "T12Line_Item," & "T12Date," & "T12Time," & "T12Stock_Code," & "T12Qty," & "T12Status," & "T12Confirm_By) " & "values(" & Delivary_Ref & ",'" & strSales_Order & "'," & strLine_Item & ",'" & Today & "','" & Now & "','" & Trim(UltraGrid1.Rows(i).Cells(2).Value) & "','" & Trim(UltraGrid1.Rows(i).Cells(5).Value) & "','N','-')"
                        up_GetSetConsume_Grige(ncQryType, nvcFieldList1, vcWhere, connection, transaction)
                    End If
                End If


                i = i + 1
            Next

            nvcFieldList1 = "UType='PROCU'"
            dsUser = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "USR"), New SqlParameter("@vcWhereClause1", nvcFieldList1))
            If isValidDataset(dsUser) Then
                _strEmail = Trim(dsUser.Tables(0).Rows(0)("email"))
                _strUserName = Trim(dsUser.Tables(0).Rows(0)("Name"))
            End If
            transaction.Commit()
            Call Update_Records_DYEDYARN_REQUEST()
            Call Send_Email_To_Projection(_strUserName, _strEmail)
            Cursor = Cursors.Default
            Me.Close()

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)

                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
                connection.Close()
            End If
        End Try
    End Sub

    Function Load_Dye_YarnCombo()
        Dim i As Integer
        Dim _Balance As Double
        Dim dt As DataTable = New DataTable()
        ' dt.Columns.Add("ID", GetType(Integer))
        dt.Columns.Add("##", GetType(String))
        For Each uRow As UltraGridRow In dg1.Rows
            _Balance = 0
            If Trim(dg1.Rows(i).Cells(0).Text) <> "" Then
                If Trim(dg1.Rows(i).Cells(9).Text) <> "" Then
                    _Balance = Trim(dg1.Rows(i).Cells(7).Value) - Trim(dg1.Rows(i).Cells(9).Text)
                    If _Balance > 0 Then
                       

                        dt.Rows.Add(New Object() {Trim(dg1.Rows(i).Cells(1).Text)})
                        
                        dt.AcceptChanges()

                        
                    End If
                Else
                    dt.Rows.Add(New Object() {Trim(dg1.Rows(i).Cells(1).Text)})

                    dt.AcceptChanges()
                End If
            End If
            i = i + 1
        Next

        Me.cboDyed_Yarn.SetDataBinding(dt, Nothing)
        cboDyed_Yarn.DisplayMember = "##"
        cboDyed_Yarn.Rows.Band.Columns(0).Width = 350
        '  Me.UltraDropDown1.ValueMember = "ID"

    End Function
   

    Private Sub cboDyed_Yarn_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboDyed_Yarn.KeyUp
        If e.KeyCode = 13 Then
            txtDy_Week.Focus()
        ElseIf e.KeyCode = Keys.Tab Then
            txtDy_Week.Focus()
        End If
    End Sub

    Private Sub txtDy_Week_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDy_Week.KeyUp
        If e.KeyCode = 13 Then
            txtDy_Year.Focus()
        ElseIf e.KeyCode = Keys.Tab Then
            txtDy_Year.Focus()
        End If
    End Sub

    Private Sub txtDy_Year_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDy_Year.KeyUp
        If e.KeyCode = 13 Then
            txtDy_Capacity.Focus()
        ElseIf e.KeyCode = Keys.Tab Then
            txtDy_Capacity.Focus()
        End If
    End Sub

    Private Sub txtDy_Capacity_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDy_Capacity.KeyUp
        On Error Resume Next
        Dim Value As Double
        Dim _qty As String

        If e.KeyCode = 13 Then
            If cboDyed_Yarn.Text <> "" Then
            Else
                MsgBox("Please enter the Yarn Name", MsgBoxStyle.Information, "Information .....")
                Exit Sub
            End If

            If txtDy_Week.Text <> "" Then
                If IsNumeric(txtDy_Week.Text) Then
                Else
                    MsgBox("Please enter the Week", MsgBoxStyle.Information, "Information .....")
                    Exit Sub
                End If
            Else
                MsgBox("Please enter the Week", MsgBoxStyle.Information, "Information .....")
                Exit Sub
            End If

            If txtDy_Year.Text <> "" Then
                If IsNumeric(txtDy_Year.Text) Then
                Else
                    MsgBox("Please enter the Year", MsgBoxStyle.Information, "Information .....")
                    Exit Sub
                End If
            Else
                MsgBox("Please enter the Year", MsgBoxStyle.Information, "Information .....")
                Exit Sub
            End If


            If txtDy_Capacity.Text <> "" Then
                If IsNumeric(txtDy_Capacity.Text) Then
                Else
                    MsgBox("Please enter the correct Capacity", MsgBoxStyle.Information, "Information .....")
                    Exit Sub
                End If
            Else
                MsgBox("Please enter the Capacity", MsgBoxStyle.Information, "Information .....")
                Exit Sub
            End If

            Dim newRow As DataRow = c_dataCustomer_DyeYarn.NewRow
            newRow("Yarn Name") = cboDyed_Yarn.Text
            newRow("Week") = txtDy_Week.Text
            newRow("Year") = txtDy_Year.Text
            Value = txtDy_Capacity.Text
            _Qty = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _Qty = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow("Capacity") = _Qty

            c_dataCustomer_DyeYarn.Rows.Add(newRow)

            txtDy_Capacity.Text = ""
            txtDy_Week.Text = ""
            txtDy_Year.Text = Year(Today)
            cboDyed_Yarn.Text = ""
            cboDyed_Yarn.ToggleDropdown()

        End If
    End Sub


    Function Send_Email_To_Projection(ByVal strProcument As String, ByVal strEmail As String)
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        ' con1 = DBEngin1.GetConnection(True)
        Dim M01 As DataSet
        Try
            Dim exc1 As New Microsoft.Office.Interop.Excel.Application
            Dim objApp As Object
            Dim objEmail As Object
           

            Dim workbooks1 As Microsoft.Office.Interop.Excel.Workbooks = exc1.Workbooks
            Dim workbook As Microsoft.Office.Interop.Excel._Workbook = workbooks1.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet)
            Dim sheets As Microsoft.Office.Interop.Excel.Sheets = workbook.Worksheets
            Dim worksheet1 As Microsoft.Office.Interop.Excel._Worksheet = CType(sheets.Item(1), Microsoft.Office.Interop.Excel._Worksheet)
            Dim range1 As Microsoft.Office.Interop.Excel.Range

            Dim A As String
            Dim I As Integer
            Dim x As Integer
            Dim Z As Integer

            objApp = CreateObject("Outlook.Application")
            objEmail = objApp.CreateItem(0)


            With objEmail

                .To = strEmail
                .Subject = "Yarn Request for /" & Trim(strSales_Order) & "-" & strLine_Item
                exc1.Visible = False

                'Dim sheetsM As Microsoft.Office.Interop.Excel.Sheets = workbook.Worksheets
                'Dim worksheet_Pro As Microsoft.Office.Interop.Excel._Worksheet = CType(sheetsM.Item(1), Microsoft.Office.Interop.Excel._Worksheet)
                worksheet1.Rows(2).Font.size = 12
                worksheet1.Rows(2).Font.Bold = True
                worksheet1.Rows(2).rowheight = 50
                With worksheet1
                    .Cells(2, 1) = "Special Yarn"
                    .Cells(2, 1).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                    ' worksheet1.Columns("A").ColumnWidth = 12
                    .Range("A2:H2").MergeCells = True
                    .Range("A2:H2").VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
                    .Range("A2:H2").Interior.Color = RGB(0, 112, 192)

                    .Columns("A").ColumnWidth = 10
                    .Columns("B").ColumnWidth = 10
                    .Columns("C").ColumnWidth = 55
                    .Columns("D").ColumnWidth = 10
                    .Columns("E").ColumnWidth = 10
                    .Columns("F").ColumnWidth = 10
                    .Columns("G").ColumnWidth = 10
                    .Columns("H").ColumnWidth = 18

                End With

                A = 97
                I = 0
                For I = 1 To 8
                    With worksheet1
                        .Range(Chr(A) & "2", Chr(A) & "2").Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                        .Range(Chr(A) & "2", Chr(A) & "2").Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                        .Range(Chr(A) & "2", Chr(A) & "2").Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                        .Range(Chr(A) & "2", Chr(A) & "2").Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous

                    End With
                    A = A + 1
                Next
                x = 3
                worksheet1.Rows(x).Font.size = 10
                worksheet1.Rows(x).Font.Bold = True
                worksheet1.Rows(x).rowheight = 30
                With worksheet1
                    .Cells(x, 1) = "Quality"
                    .Cells(x, 1).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter

                    .Cells(x, 2) = "10 Class"
                    .Cells(x, 2).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter

                    .Cells(x, 3) = "Yarn Description"
                    .Cells(x, 3).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter

                    .Cells(x, 4) = "Qty(Kg)"
                    .Cells(x, 4).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter

                    .Cells(x, 5) = "Sales Order"
                    .Cells(x, 5).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter

                    .Cells(x, 6) = "Line Item"
                    .Cells(x, 6).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter

                    .Cells(x, 7) = "Capacity"
                    .Cells(x, 7).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter

                    .Cells(x, 7).WrapText = True
                    .Cells(x, 8) = "Weekly Consumption"
                    .Cells(x, 8).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter

                    '.Cells(x, 9) = "Weekly Consumption"
                    '.Cells(x, 9).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter

                End With

                A = 97
                I = 0
                For I = 1 To 8
                    With worksheet1
                        .Range(Chr(A) & "3", Chr(A) & "3").Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                        .Range(Chr(A) & "3", Chr(A) & "3").Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                        .Range(Chr(A) & "3", Chr(A) & "3").Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                        .Range(Chr(A) & "3", Chr(A) & "3").Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                        .Range(Chr(A) & "3", Chr(A) & "3").Interior.Color = RGB(169, 208, 142)

                        .Range(Chr(A) & "3", Chr(A) & "3").MergeCells = True
                        .Range(Chr(A) & "3", Chr(A) & "3").VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
                    End With
                    A = A + 1
                Next

                Sql = "T14Sales_order='" & strSales_Order & "' AND T14Line_Item=" & strLine_Item & " AND T14Ref_no=" & Delivary_Ref & ""
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetYarn_Booking", New SqlParameter("@cQryType", "YRNR"), New SqlParameter("@vcWhereClause1", Sql))
                Z = 0
                x = x + 1
                For Each DTRow3 As DataRow In M01.Tables(0).Rows
                    worksheet1.Rows(x).Font.size = 10
                    worksheet1.Rows(x).rowheight = 20
                    With worksheet1
                        .Cells(x, 1) = txtQuality.Text
                        .Cells(x, 1).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter

                        .Cells(x, 2) = M01.Tables(0).Rows(Z)("T14Class")
                        .Cells(x, 2).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter


                        .Cells(x, 3) = M01.Tables(0).Rows(Z)("T14Yarn")
                        .Cells(x, 3).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter


                        .Cells(x, 4) = M01.Tables(0).Rows(Z)("T14Available")
                        .Cells(x, 4).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                        range1 = .Cells(x, 4)
                        range1.NumberFormat = "0.00"

                        .Cells(x, 5) = M01.Tables(0).Rows(Z)("T14Sales_order")
                        .Cells(x, 5).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter

                        .Cells(x, 6) = M01.Tables(0).Rows(Z)("T14Line_Item")
                        .Cells(x, 6).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter


                        .Cells(x, 7) = "Week " & " " & Trim(M01.Tables(0).Rows(Z)("T14Week"))
                        .Cells(x, 7).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter

                        .Cells(x, 8) = M01.Tables(0).Rows(Z)("T14Capacity")
                        .Cells(x, 8).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                        range1 = .Cells(x, 8)
                        range1.NumberFormat = "0.00"


                        A = 97
                        I = 0
                        For I = 1 To 8
                            With worksheet1
                                .Range(Chr(A) & x, Chr(A) & x).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                                .Range(Chr(A) & x, Chr(A) & x).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                                .Range(Chr(A) & x, Chr(A) & x).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                                .Range(Chr(A) & x, Chr(A) & x).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                                '.Range(Chr(A) & "3", Chr(A) & "3").Interior.Color = RGB(169, 208, 142)

                                .Range(Chr(A) & x, Chr(A) & x).MergeCells = True
                                .Range(Chr(A) & x, Chr(A) & x).VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
                            End With
                            A = A + 1
                        Next
                    End With

                    x = x + 1
                    Z = Z + 1
                Next




                Dim xlRn As Microsoft.Office.Interop.Excel.Range
                Dim Connect As String
                Dim strbody As String

                'strBody = "This is a test " & vbCrLf & vbCrLf & "Thanks Michael"
                '  RangetoHTML(xlRn)

                Connect = worksheet1.Range("A2:H" & x - 1).Copy()
                xlRn = worksheet1.Range("A2:H" & x + 1)
                'xlRn.Copy()

                '.HTMLBody = " Dear" & Trim(cboPlaner.Text) & "," & vbNewLine & _
                '              "Please Quote best possible delivery for below" & Chr(10) _
                '                                   & RangetoHTML(xlRn)

                .HTMLBody = "Dear " & Trim(strProcument) & ",<br>Please order below yarns " & RangetoHTML(xlRn)

                .display()

            End With
            objEmail = Nothing
            objApp = Nothing

            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            con.close()
            ' worksheet1.Cells(4, 5) = _Fail_Batch
            'worksheet1.Cells(4, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
            ' MsgBox("Report Genarated successfully", MsgBoxStyle.Information, "Technova ....")
            ' MsgBox(_Fail_Batch)
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
                con.close()
            End If
        End Try
    End Function


    Function RangetoHTML(ByVal rng As Microsoft.Office.Interop.Excel.Range)
        ' Changed by Ron de Bruin 28-Oct-2006
        ' Working in Office 2000-2010
        Dim fso As Object
        Dim ts As Object
        Dim TempFile As String
        ' Dim TempWB As Microsoft.Office.Interop.Excel.Workbook

        Dim exc As New Microsoft.Office.Interop.Excel.Application
        Dim TempWB1 As Microsoft.Office.Interop.Excel.Workbooks = exc.Workbooks
        Dim TempWB As Microsoft.Office.Interop.Excel._Workbook = TempWB1.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet)

        TempFile = Environ$("temp") & "/" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

        'Copy the range and create a new workbook to past the data in
        rng.Copy()
        'TempWB = Microsoft.Office.Interop.Excel.Workbooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet)




        With TempWB.Sheets(1)
            .Cells(1).PasteSpecial(Paste:=8)
            ' Microsoft.Office.Interop.Excel.XlPastef
            '.Cells(1).PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues, , False, False)
            '.Cells(1).PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats, , False, False)
            '.Cells(1).Select()
            'Application.CutCopyMode = False
            On Error Resume Next
            .DrawingObjects.Visible = True
            .DrawingObjects.Delete()
            On Error GoTo 0
        End With


        'Publish the sheet to a htm file
        With TempWB.PublishObjects.Add( _
             SourceType:=Microsoft.Office.Interop.Excel.XlSourceType.xlSourceRange, _
             Filename:=TempFile, _
             Sheet:=TempWB.Sheets(1).Name, _
             Source:=TempWB.Sheets(1).UsedRange.Address, _
             HtmlType:=Microsoft.Office.Interop.Excel.XlHtmlType.xlHtmlStatic)
            .Publish(True)
        End With

        'Read all data from the htm file into RangetoHTML
        fso = CreateObject("Scripting.FileSystemObject")
        ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
        RangetoHTML = ts.ReadAll
        ts.Close()
        RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                              "align=left x:publishsource=")

        'Close TempWB
        TempWB.Close(savechanges:=False)

        'Delete the htm file we used in this function
        Kill(TempFile)

        ts = Nothing
        fso = Nothing
        TempWB = Nothing
    End Function


    Private Sub cboY_Name_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboY_Name.KeyUp
        If e.KeyCode = 13 Then
            txtY_Week.Focus()
        ElseIf e.KeyCode = Keys.Tab Then
            txtY_Week.Focus()
        End If
    End Sub

    Private Sub txtY_Year_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtY_Year.KeyUp
        If e.KeyCode = 13 Then
            txtY_Qty.Focus()
        ElseIf e.KeyCode = Keys.Tab Then
            txtY_Qty.Focus()
        End If
    End Sub

    Private Sub txtY_Week_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtY_Week.KeyUp
        If e.KeyCode = 13 Then
            txtY_Year.Focus()
        ElseIf e.KeyCode = Keys.Tab Then
            txtY_Year.Focus()
        End If
    End Sub

    Private Sub txtY_Qty_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtY_Qty.KeyUp
        On Error Resume Next
        Dim Value As Double
        Dim _qty As String

        If e.KeyCode = 13 Then
            If cboY_Name.Text <> "" Then
            Else
                MsgBox("Please enter the Yarn Name", MsgBoxStyle.Information, "Information .....")
                Exit Sub
            End If

            If txtY_Week.Text <> "" Then
                If IsNumeric(txtY_Week.Text) Then
                Else
                    MsgBox("Please enter the Week", MsgBoxStyle.Information, "Information .....")
                    Exit Sub
                End If
            Else
                MsgBox("Please enter the Week", MsgBoxStyle.Information, "Information .....")
                Exit Sub
            End If

            If txtY_Year.Text <> "" Then
                If IsNumeric(txtY_Year.Text) Then
                Else
                    MsgBox("Please enter the Year", MsgBoxStyle.Information, "Information .....")
                    Exit Sub
                End If
            Else
                MsgBox("Please enter the Year", MsgBoxStyle.Information, "Information .....")
                Exit Sub
            End If


            If txtY_Qty.Text <> "" Then
                If IsNumeric(txtY_Qty.Text) Then
                Else
                    MsgBox("Please enter the correct Capacity", MsgBoxStyle.Information, "Information .....")
                    Exit Sub
                End If
            Else
                MsgBox("Please enter the Capacity", MsgBoxStyle.Information, "Information .....")
                Exit Sub
            End If

            Dim newRow As DataRow = c_dataCustomer_Yarn.NewRow
            newRow("Yarn Name") = cboY_Name.Text
            newRow("Week") = txtY_Week.Text
            newRow("Year") = txtY_Year.Text
            Value = txtY_Qty.Text
            _qty = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _qty = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow("Capacity") = _qty

            c_dataCustomer_Yarn.Rows.Add(newRow)

            txtY_Qty.Text = ""
            txtY_Week.Text = ""
            txtY_Year.Text = Year(Today)
            cboY_Name.Text = ""
            cboY_Name.ToggleDropdown()
        End If
    End Sub

    Private Sub UltraButton28_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton28.Click
        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True


        Dim vcWhere As String
        Dim M01 As DataSet
        Dim ncQryType As String
        Dim i As Integer
        Dim _strUserName As String
        Dim _strEmail As String

        i = 0

        Try
            Cursor = Cursors.WaitCursor
            For Each uRow As UltraGridRow In UltraGrid1.Rows
                ' MsgBox(UltraGrid1.Rows(i).Cells(5).Value)
                If IsNumeric(UltraGrid1.Rows(i).Cells(5).Value) Then
                    vcWhere = "T12Ref_No=" & Delivary_Ref & " and T12Sales_Order='" & strSales_Order & "' and T12Line_Item=" & strLine_Item & " and T12Stock_Code='" & Trim(UltraGrid1.Rows(i).Cells(2).Value) & "'"
                    M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetUse_Griege_Qty", New SqlParameter("@cQryType", "CGS1"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M01) Then
                    Else
                        ncQryType = "GADD"
                        nvcFieldList1 = "(T12Ref_No," & "T12Sales_Order," & "T12Line_Item," & "T12Date," & "T12Time," & "T12Stock_Code," & "T12Qty," & "T12Status," & "T12Confirm_By) " & "values(" & Delivary_Ref & ",'" & strSales_Order & "'," & strLine_Item & ",'" & Today & "','" & Now & "','" & Trim(UltraGrid1.Rows(i).Cells(2).Value) & "','" & Trim(UltraGrid1.Rows(i).Cells(5).Value) & "','N','-')"
                        up_GetSetConsume_Grige(ncQryType, nvcFieldList1, vcWhere, connection, transaction)
                    End If
                End If


                i = i + 1
            Next

            nvcFieldList1 = "UType='PROCU'"
            dsUser = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "USR"), New SqlParameter("@vcWhereClause1", nvcFieldList1))
            If isValidDataset(dsUser) Then
                _strEmail = Trim(dsUser.Tables(0).Rows(0)("email"))
                _strUserName = Trim(dsUser.Tables(0).Rows(0)("Name"))
            End If
            transaction.Commit()
            Call Update_Records_YARN_REQUEST_PROCUMENT()
            Call Send_Email_To_Projection(_strUserName, _strEmail)
            Cursor = Cursors.Default
            Me.Close()

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)

                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
                connection.Close()
            End If
        End Try
    End Sub

    Function Load_Dye_Mc()
        Dim Sql As String
        Dim M01 As DataSet
        Dim M012 As DataSet
        Dim i As Integer

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim vcWhere As String
        Try

            Sql = "select * from M16Quality_RCode where M16Material='" & _base30class & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                Sql = "select * from M14R_Code where M14Order='" & Trim(M01.Tables(0).Rows(0)("M16R_Code")) & "'"
                M012 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(M012) Then
                    If dg_Dye_Detailes.Rows(0).Cells.Count >= 2 Then
                        For i = 1 To dg_Dye_Detailes.Rows(0).Cells.Count - 1
                            
                        Next
                    End If
                End If

            End If

            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function
    Function Load_Dye_MC_Group()
        Dim agroup1 As UltraGridGroup
        Dim agroup3 As UltraGridGroup
        Dim _MCGroup As String
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim M02 As DataSet

        Dim vcwhere As String
        Dim vcwhere1 As String
        Dim i As Integer
        Dim _LineItem As String
        Dim _Week As String
        Dim _Week1 As String
        Dim _Rowcount As Integer
        Dim _ColumCount As Integer
        Dim _ColumCount1 As Integer
        Dim _WeekNo As Integer
        Dim T01 As DataSet
        Dim T02 As DataSet
        Dim tmpRow As Integer
        Dim _Lane As String
        Dim _1stWeek As Integer

        Dim Z As Integer
        dg_Dye_MC.DisplayLayout.Bands(0).Groups.Clear()
        dg_Dye_MC.DisplayLayout.Bands(0).Columns.Dispose()
        Try
            agroup1 = dg_Dye_MC.DisplayLayout.Bands(0).Groups.Add("Machine No")
            agroup1.Width = 80
            Dim dt As DataTable = New DataTable()
            ' dt.Columns.Add("ID", GetType(Integer))
            Dim colWork As New DataColumn("##", GetType(String))
            dt.Columns.Add(colWork)
            colWork.ReadOnly = False


            If chkE1.Checked = True Then
                _MCGroup = "ECO +"
            ElseIf chkE2.Checked = True Then
                _MCGroup = "ECO +"
            ElseIf chkEco1.Checked = True Then
                _MCGroup = "ECO"
            ElseIf chkEco2.Checked = True Then
                _MCGroup = "ECO"
            ElseIf chkLR1.Checked = True Then
                _MCGroup = "LR"
            ElseIf chkLR2.Checked = True Then
                _MCGroup = "LR"
            ElseIf chkThan1.Checked = True Then
                _MCGroup = "THAN"
            ElseIf chkTHAN2.Checked = True Then
                _MCGroup = "THAN"
            End If
            _ColumCount = 3
            _Lane = ""
            _Rowcount = 0
            For i = 0 To dg_Dye_Shade.DisplayLayout.Bands(0).Columns.Count - 1
                Dim _Status As Boolean

                Z = 0
                _Status = False
                For Z = 0 To dg_Dye_Shade.Rows.Count - 5
                    '  MsgBox(dg_Dye_Shade.Rows(Z).Cells(_ColumCount).Value)
                    If dg_Dye_Shade.Rows(Z).Cells(_ColumCount).Value = True Then
                        If _Lane <> "" Then
                            _Lane = _Lane & "','" & Microsoft.VisualBasic.Left(dg_Dye_Shade.Rows(Z).Cells(0).Text, 1)
                            _Status = True
                            _ColumCount = _ColumCount + 2
                            Exit For
                        Else
                            _Lane = Microsoft.VisualBasic.Left(dg_Dye_Shade.Rows(Z).Cells(0).Text, 1)
                            _Status = True
                            _ColumCount = _ColumCount + 3
                            Exit For
                        End If
                    End If
                Next
                If _Status = False Then
                    _ColumCount = _ColumCount + 2
                End If
                ' MsgBox(dg_Dye_Shade.DisplayLayout.Bands(0).Columns.Count)
                If dg_Dye_Shade.DisplayLayout.Bands(0).Columns.Count <= _ColumCount Then
                    Exit For
                End If
            Next
            vcwhere = "M50MC_Group='" & _MCGroup & "' and m50Lane in ('" & _Lane & "')"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "LODD"), New SqlParameter("@vcWhereClause1", vcwhere))
            i = 0
            _Dye_Noof_Lane = M01.Tables(0).Rows.Count
            For Each DTRow5 As DataRow In M01.Tables(0).Rows
                dt.Rows.Add(M01.Tables(0).Rows(i)("M50MC_No"))
                i = i + 1
            Next
            'tmpRow = i + 1
            'i = 0

            'For i = 0 To 4
            '    dt.Rows.Add("")
            'Next

            ''tmpRow = i + 1
            Me.dg_Dye_MC.SetDataBinding(dt, Nothing)
            Me.dg_Dye_MC.DisplayLayout.Bands(0).Columns(0).Group = agroup1
            '========================================================================

            For i = 1 To dg_Dye_Detailes.DisplayLayout.Bands(0).Columns.Count - 1
                Dim _St As String

                If IsNumeric(dg_Dye_Detailes.Rows(4).Cells(i).Value) Then
                    _St = _WeekNo & "-A"
                    'INSERT TABLE
                    _WeekNo = Microsoft.VisualBasic.Right(dg_Dye_Detailes.DisplayLayout.Bands(0).Columns(i).Header.Caption, 2)

                    _Week = "Week " & _WeekNo
                    _Week1 = "Week- " & _WeekNo

                    agroup3 = dg_Dye_MC.DisplayLayout.Bands(0).Groups.Add(_Week)
                    agroup3.Header.Caption = _Week
                    Dim Z2 As Integer
                    _Week1 = "Week- " & _WeekNo & "A"
                    Me.dg_Dye_MC.DisplayLayout.Bands(0).Columns.Add(_St, "Start Time")
                    Me.dg_Dye_MC.DisplayLayout.Bands(0).Columns(_St).Group = agroup3
                    Me.dg_Dye_MC.DisplayLayout.Bands(0).Columns(_St).Width = 70
                    Me.dg_Dye_MC.DisplayLayout.Bands(0).Columns(_St).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

                    _St = _WeekNo & "-B"
                    Me.dg_Dye_MC.DisplayLayout.Bands(0).Columns.Add(_St, "End Time")
                    Me.dg_Dye_MC.DisplayLayout.Bands(0).Columns(_St).Group = agroup3
                    Me.dg_Dye_MC.DisplayLayout.Bands(0).Columns(_St).Width = 70
                    Me.dg_Dye_MC.DisplayLayout.Bands(0).Columns(_St).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

                    _St = _WeekNo & "-C"
                    Me.dg_Dye_MC.DisplayLayout.Bands(0).Columns.Add(_St, "Batch")
                    Me.dg_Dye_MC.DisplayLayout.Bands(0).Columns(_St).Group = agroup3
                    Me.dg_Dye_MC.DisplayLayout.Bands(0).Columns(_St).Width = 70
                    Me.dg_Dye_MC.DisplayLayout.Bands(0).Columns(_St).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

                    _St = _WeekNo & "-D"
                    Me.dg_Dye_MC.DisplayLayout.Bands(0).Columns.Add(_St, "##")
                    Me.dg_Dye_MC.DisplayLayout.Bands(0).Columns(_St).Group = agroup3
                    Me.dg_Dye_MC.DisplayLayout.Bands(0).Columns(_St).Width = 70
                    Me.dg_Dye_MC.DisplayLayout.Bands(0).Columns(_St).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                End If

            Next
            _ColumCount = dg_Dye_MC.DisplayLayout.Bands(0).Columns.Count - 1
            _ColumCount = _ColumCount / 4
            If _ColumCount = 0 Then
                _ColumCount = 1
            End If
            For i = 0 To dg_Dye_MC.Rows.Count - 1
                _ColumCount1 = 4
                For Z = 1 To _ColumCount
                    dg_Dye_MC.Rows(i).Cells(_ColumCount1).Style = ColumnStyle.CheckBox
                    dg_Dye_MC.Rows(i).Cells(_ColumCount1).Value = False
                    _ColumCount1 = _ColumCount1 + 4
                Next
            Next

            '===========================================================================================================================
            Z = 0
            Dim weekStart As DateTime
            Dim _BatchNo As Integer
            Dim _QtyMin As Integer
            Dim _MCLane As Integer

            _ColumCount = 5
            _ColumCount1 = 1
            For i = 1 To dg_Dye_Shade.DisplayLayout.Bands(0).Columns.Count - 1
                If i = 1 Then
                    ' MsgBox(dg_Dye_Detailes.Rows(4).Cells(i).Text)
                    _1stWeek = Microsoft.VisualBasic.Right(dg_Dye_Detailes.DisplayLayout.Bands(0).Columns(i).Header.Caption, 2)
                End If
                _WeekNo = Microsoft.VisualBasic.Right(dg_Dye_Detailes.DisplayLayout.Bands(0).Columns(i).Header.Caption, 2)
                Z = 0
                For Z = 0 To dg_Dye_MC.Rows.Count - 1
                    If _WeekNo >= _1stWeek Then
                        vcwhere = "tmpMachine='" & Trim(dg_Dye_MC.Rows(0).Cells(0).Text) & "' and tmpWeek=" & _WeekNo & " and tmpYear='" & Year(Today) & "'"
                        M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "DYB"), New SqlParameter("@vcWhereClause1", vcwhere))
                        If isValidDataset(M01) Then
                            If IsNumeric(dg_Dye_Shade.Rows(_Rowindex - 3).Cells(_ColumCount1).Text) Then
                                _BatchNo = dg_Dye_Shade.Rows(_Rowindex - 3).Cells(_ColumCount1).Text
                            End If

                        Else
                            weekStart = GetWeekStartDate(_WeekNo, 2015)
                            weekStart = weekStart.AddDays(-4)
                            weekStart = weekStart & " " & "7:30AM"
                            dg_Dye_MC.Rows(Z).Cells(_ColumCount - 4).Value = weekStart
                            _Rowindex = dg_Dye_Shade.Rows.Count
                            If IsNumeric(dg_Dye_Shade.Rows(_Rowindex - 3).Cells(_ColumCount1).Text) Then
                                _BatchNo = dg_Dye_Shade.Rows(_Rowindex - 3).Cells(_ColumCount1).Text
                            End If

                            vcwhere = "M50Mc_no='" & Trim(dg_Dye_MC.Rows(0).Cells(0).Text) & "'"
                            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "LODD"), New SqlParameter("@vcWhereClause1", vcwhere))
                            If isValidDataset(M01) Then
                                _MCLane = M01.Tables(0).Rows(0)("M50Lane")
                            End If

                            vcwhere = "M48MC_Group='" & _MCGroup & "' and M48Lane=" & _MCLane & ""
                            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "DMCC1"), New SqlParameter("@vcWhereClause1", vcwhere))
                            If isValidDataset(M01) Then
                                _QtyMin = M01.Tables(0).Rows(0)("M48Total_HR") * _BatchNo
                            End If
                            weekStart = weekStart.AddMinutes(+_QtyMin)
                            dg_Dye_MC.Rows(Z).Cells(_ColumCount - 3).Value = weekStart
                            dg_Dye_MC.Rows(Z).Cells(_ColumCount - 2).Value = _BatchNo
                            _ColumCount = _ColumCount + 4
                            ' MsgBox(dg_Dye_MC.DisplayLayout.Bands(0).Columns.Count - 1)
                            If (dg_Dye_MC.DisplayLayout.Bands(0).Columns.Count) >= _ColumCount Then
                                _ColumCount1 = _ColumCount1 + 3
                                Continue For
                            Else
                                con.close()
                                Exit Function
                            End If
                        End If
                    Else
                        vcwhere = "tmpMachine='" & Trim(dg_Dye_MC.Rows(Z).Cells(0).Text) & "' and tmpWeek=" & _WeekNo & " and tmpYear='" & Year(Today) + 1 & "'"
                        M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "DYB"), New SqlParameter("@vcWhereClause1", vcwhere))
                    End If
                Next
            Next
            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Private Sub cmdD_Save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdD_Save.Click
        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Dim vcWhere As String
        Dim ncQryType As String
        Dim _MCGroup As String
        Dim i As Integer
        Dim _WeekNo As Integer
        Dim _1STWEEK As Integer
        Dim _YEAR As Integer
        Dim _LeadTime As String

        Try
            'If chkNormal.Checked = True Then
            '    _LeadTime = "Normal"
            'ElseIf chkShort.Checked = True Then
            '    _LeadTime = "Short"
            'End If

            'If _LeadTime <> "" Then
            'Else
            '    MsgBox("Please select the Lead time", MsgBoxStyle.Information, "Information ....")
            '    Exit Sub
            'End If

            'If UltraGrid3.Rows.Count > 0 Then
            'Else
            '    MsgBox("Please enter the delivery plan", MsgBoxStyle.Information, "Information ......")
            '    Exit Sub
            'End If
            If chkE1.Checked = True Then
                _MCGroup = "ECO +"
            ElseIf chkE2.Checked = True Then
                _MCGroup = "ECO +"
            ElseIf chkEco1.Checked = True Then
                _MCGroup = "ECO"
            ElseIf chkEco2.Checked = True Then
                _MCGroup = "ECO"
            ElseIf chkLR1.Checked = True Then
                _MCGroup = "LR"
            ElseIf chkLR2.Checked = True Then
                _MCGroup = "LR"
            ElseIf chkThan1.Checked = True Then
                _MCGroup = "THAN"
            ElseIf chkTHAN2.Checked = True Then
                _MCGroup = "THAN"
            End If

            nvcFieldList1 = "DELETE FROM T18Delivary_Plane WHERE T18Ref_No=" & Delivary_Ref & " and T18Sales_Order='" & strSales_Order & "' and T18Line_Item in ('" & _LineItem & "')"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            Dim _Colume As Integer
            Dim Y As Integer
            Dim Z As Integer
            Dim _Lane As String
            _Colume = 3
            For i = 1 To dg_Dye_Detailes.DisplayLayout.Bands(0).Columns.Count - 1
                Dim _St As String

                If IsNumeric(dg_Dye_Detailes.Rows(4).Cells(i).Value) Then

                    'INSERT TABLE
                    _WeekNo = Microsoft.VisualBasic.Right(dg_Dye_Detailes.DisplayLayout.Bands(0).Columns(i).Header.Caption, 2)
                    If i = 1 Then
                        _1STWEEK = _WeekNo
                    End If

                    If _WeekNo > _1STWEEK Then
                        _YEAR = Year(Today)
                    Else
                        _YEAR = Year(Today) - 1
                    End If

                    Y = 0
                    For Each uRow As UltraGridRow In dg_Dye_Shade.Rows

                        If dg_Dye_Shade.Rows(Y).Cells(_Colume).Value = True Then
                            _Lane = dg_Dye_Shade.Rows(Y).Cells(0).Value
                            Exit For
                        End If
                        Y = Y + 1
                    Next
                    _Colume = _Colume + 3
                    ncQryType = "DLIP"
                    nvcFieldList1 = "(T18Ref_No," & "T18Sales_Order," & "T18Line_Item," & "T18WeekNo," & "T18Year," & "T18Qty," & "T18Sub," & "T18App," & "T18MC," & "T18Lead_Time," & "T18Lane) " & "values(" & Delivary_Ref & ",'" & strSales_Order & "'," & _LineItem & "," & _WeekNo & "," & _YEAR & ",'" & dg_Dye_Detailes.Rows(4).Cells(i).Value & "','N','N','" & _MCGroup & "','" & _LeadTime & "','" & _Lane & "')"
                    up_GetSetCAPACITY(ncQryType, nvcFieldList1, vcWhere, connection, transaction)
                ElseIf Trim(dg_Dye_Detailes.Rows(4).Cells(i).Text) = "SUB" Then
                    _WeekNo = Microsoft.VisualBasic.Right(dg_Dye_Detailes.DisplayLayout.Bands(0).Columns(i).Header.Caption, 2)
                    If i = 1 Then
                        _1STWEEK = _WeekNo
                    End If

                    If _WeekNo > _1STWEEK Then
                        _YEAR = Year(Today)
                    Else
                        _YEAR = Year(Today) - 1
                    End If
                    _Lane = " "
                    ncQryType = "DLIP"
                    nvcFieldList1 = "(T18Ref_No," & "T18Sales_Order," & "T18Line_Item," & "T18WeekNo," & "T18Year," & "T18Qty," & "T18Sub," & "T18App," & "T18MC," & "T18Lead_Time," & "T18Lane) " & "values(" & Delivary_Ref & ",'" & strSales_Order & "'," & _LineItem & "," & _WeekNo & "," & _YEAR & ",'0','Y','N','" & _MCGroup & "','" & _LeadTime & "','" & _Lane & "')"
                    up_GetSetCAPACITY(ncQryType, nvcFieldList1, vcWhere, connection, transaction)
                ElseIf Trim(dg_Dye_Detailes.Rows(4).Cells(i).Text) = "APP" Then
                    _WeekNo = Microsoft.VisualBasic.Right(dg_Dye_Detailes.DisplayLayout.Bands(0).Columns(i).Header.Caption, 2)
                    If i = 1 Then
                        _1STWEEK = _WeekNo
                    End If

                    If _WeekNo > _1STWEEK Then
                        _YEAR = Year(Today)
                    Else
                        _YEAR = Year(Today) - 1
                    End If
                    _Lane = " "
                    ncQryType = "DLIP"
                    nvcFieldList1 = "(T18Ref_No," & "T18Sales_Order," & "T18Line_Item," & "T18WeekNo," & "T18Year," & "T18Qty," & "T18Sub," & "T18App," & "T18MC," & "T18Lead_Time," & "T18Lane) " & "values(" & Delivary_Ref & ",'" & strSales_Order & "'," & _LineItem & "," & _WeekNo & "," & _YEAR & ",'0','N','Y','" & _MCGroup & "','" & _LeadTime & "','" & _Lane & "')"
                    up_GetSetCAPACITY(ncQryType, nvcFieldList1, vcWhere, connection, transaction)

                End If
                ' _Colume = _Colume + 3
            Next

            nvcFieldList1 = "DELETE FROM tmpBlock_Dye_Machine WHERE tmpRef_No=" & Delivary_Ref & ""
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            transaction.Commit()
            connection.Close()
            Call Load_Gride_Dye_Grige() 'Screen 1 Gride Creation
            Call Load_Gride_Dye_Projection()

            Call Load_Gride_Dye_Main() 'Screen 1 Data Filling

            Call Load_Gride_Dye_Grige_Detailes() 'Screen 2 Creation
            Call Load_Gride_Dye_Capacity() 'Dye Capacity
            Call Dye_Shade_Gride_Main()

            Call Load_Gride_Delivary()
            Panel20.Visible = False
            Panel21.Visible = False
            UltraTabControl1.SelectedTab = UltraTabControl1.Tabs(6)
            Call Load_Grid_Delivary_Plan()

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)

                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
                connection.Close()
            End If
        End Try
    End Sub

    Private Sub chkShort_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If chkShort.Checked = True Then
            chkNormal.Checked = False
            Panel21.Visible = True
            chkS_Date.Visible = True
            chkS_Week.Visible = True
            lblDelivary_Balance.Text = lblDye_Qty.Text
        End If
    End Sub

    Private Sub chkNormal_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If chkNormal.Checked = True Then
            chkShort.Checked = False
            Panel21.Visible = True
            'chkS_Date.Visible = True
            'chkS_Week.Visible = True
            lblDelivary_Balance.Text = lblDye_Qty.Text
        End If
    End Sub

    Private Sub txtDye_Qty_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDye_Qty.KeyUp
        On Error Resume Next
        Dim Value As Double
        If e.KeyCode = 13 Then
            If IsNumeric(txtDye_Qty.Text) Then
            Else
                MsgBox("Please enter the correct Qty", MsgBoxStyle.Information, "Information ....")
                Exit Sub
            End If
            If txtDye_Qty.Text <> "" Then
            Else
                MsgBox("Please enter the correct Qty", MsgBoxStyle.Information, "Information ....")
                Exit Sub
            End If

            If IsNumeric(txtDye_Week.Text) Then
            Else
                MsgBox("Please enter the correct Week", MsgBoxStyle.Information, "Information ....")
                Exit Sub
            End If
            If txtDye_Week.Text <> "" Then
            Else
                MsgBox("Please enter the correct Week", MsgBoxStyle.Information, "Information ....")
                Exit Sub
            End If

            If txtDye_Week.Visible = True Then
                Dim newRow1 As DataRow = c_dataCustomer1_Delivary.NewRow
                newRow1("Week No") = txtDye_Week.Text
                newRow1("Qty") = txtDye_Qty.Text

                c_dataCustomer1_Delivary.Rows.Add(newRow1)
                txtDye_Week.Text = ""
                txtDye_Qty.Text = ""

            End If
        End If
    End Sub

    Private Sub chkS_Date_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkS_Date.CheckedChanged
        If chkS_Date.Checked = True Then
            lblDye_DD.Text = "Date"
            chkS_Week.Checked = False
            txtDel_Date.Visible = True
            txtDel_Date.Text = Today
            txtDye_Week.Visible = False
            Call Load_Gride_Delivary_Daily()
        End If
    End Sub

    Private Sub chkS_Week_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkS_Week.CheckedChanged
        If chkS_Week.Checked = True Then
            lblDye_DD.Text = "Week No"
            chkS_Date.Checked = False
            txtDel_Date.Visible = False
            txtDye_Week.Visible = True
            Call Load_Gride_Delivary()
        End If
    End Sub


    Private Sub cboK_Mc_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = 13 Then
            txtAllocated_Qty.Focus()
        ElseIf e.KeyCode = Keys.Tab Then
            txtAllocated_Qty.Focus()
        ElseIf e.KeyCode = Keys.Escape Then
            OPR_Kplan.Visible = False
        End If
    End Sub


    Private Sub txtAllocated_Qty_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAllocated_Qty.KeyUp
        If e.KeyCode = 13 Then

        ElseIf e.KeyCode = Keys.Tab Then

        ElseIf e.KeyCode = Keys.Escape Then
            OPR_Kplan.Visible = False
        End If
    End Sub

    Private Sub UltraGrid4_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles UltraGrid4.AfterCellUpdate
        On Error Resume Next
        Dim i As Integer
        Dim _Qty As Double

        If Check_Knt_MC() = True Then
            Exit Sub
        End If
        '   MsgBox(UltraGrid4.ActiveRow.Index)
        i = (UltraGrid4.ActiveRow.Index)
        If CDbl(lblKP_Balance.Text) > 0 Then
            If UltraGrid4.Rows(i).Cells(7).Value = True Then
                If CDbl(lblKP_Balance.Text) >= CDbl(UltraGrid4.Rows(i).Cells(3).Value) Then
                    UltraGrid4.Rows(i).Cells(4).Value = UltraGrid4.Rows(i).Cells(3).Value
                    If lblKP_Balance.Text > 0 Then
                        lblKP_Balance.Text = CDbl(lblKP_Balance.Text) - CDbl(UltraGrid4.Rows(i).Cells(3).Value)
                    End If
                End If
                '   Call Save_Temp_KnittingBord()
                _Qty = 0
                i = 0
                For Each uRow As UltraGridRow In UltraGrid4.Rows
                    If IsNumeric(UltraGrid4.Rows(i).Cells(4).Value) Then
                        _Qty = CDbl(UltraGrid4.Rows(i).Cells(4).Value) + _Qty
                    End If
                    i = i + 1
                Next

                lblKP_Balance.Text = CDbl(txtK_Qty.Text) - _Qty
            Else
                _Qty = 0
                i = 0
                For Each uRow As UltraGridRow In UltraGrid4.Rows
                    If IsNumeric(UltraGrid4.Rows(i).Cells(4).Value) Then
                        _Qty = CDbl(UltraGrid4.Rows(i).Cells(4).Value) + _Qty
                    End If
                    i = i + 1
                Next
                lblKP_Balance.Text = CDbl(txtK_Qty.Text) - _Qty
            End If
        Else
            'UltraGrid4.Rows(i).Cells(7).Value = False
            Exit Sub

        End If


    End Sub

    Function Save_Temp_KnittingBord()
        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Dim vcWhere As String
        Dim _McName As String
        Dim M01 As DataSet
        Dim ncQryType As String
        Dim _fromDate As Date
        Dim _Todate As Date

        Try
            Dim i As Integer
            If Microsoft.VisualBasic.Left(txtMC_Group_Knt.Text, 1) = "S" Then
                _McName = "SJ"
            ElseIf Microsoft.VisualBasic.Left(txtMC_Group_Knt.Text, 1) = "R" Then
                _McName = "DJ"

            End If
            '===============================================
            i = 0
            For Each uRow As UltraGridRow In UltraGrid4.Rows
                nvcFieldList1 = "tmpRef_No=" & Delivary_Ref & " and tmpMC_No='" & Trim(UltraGrid4.Rows(i).Cells(0).Value) & "' and tmpSales_Order='" & strSales_Order & "' and tmpLine_Item='" & strLine_Item & "' and tmpWeek_No=" & txtWeek_Knt.Text & " and tmpYear=" & txtYear_Knt.Text & ""
                M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "KPS"), New SqlParameter("@vcWhereClause1", nvcFieldList1))
                If isValidDataset(M01) Then
                    If UltraGrid4.Rows(i).Cells(7).Value = False Then
                        ' nvcFieldList1 = "delete from tmpKnitting_Plan_Board where tmpRef_No=" & Delivary_Ref & " and tmpMC_No='" & Trim(UltraGrid4.Rows(i).Cells(0).Value) & "' and tmpSales_Order='" & strSales_Order & "' and tmpLine_Item='" & strLine_Item & "' and tmpWeek_No=" & txtWeek_Knt.Text & " and tmpYear=" & txtYear_Knt.Text & ""
                        ' ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                    End If

                    nvcFieldList1 = "select sum(tmpKnt_Order) as qty from tmpKnitting_Plan_Board where tmpRef_No=" & Delivary_Ref & "  and tmpSales_Order='" & strSales_Order & "' and tmpLine_Item='" & strLine_Item & "' and tmpWeek_No=" & txtWeek_Knt.Text & " and tmpYear=" & txtYear_Knt.Text & " group by tmpRef_No,tmpSales_Order,tmpLine_Item,tmpWeek_No,tmpYear"
                    M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                    If isValidDataset(M01) Then

                        lblKP_Balance.Text = CDbl(lblKP_Balance.Text) + M01.Tables(0).Rows(0)("Qty")
                        'lblKP_Balance.Text = Microsoft.VisualBasic.Format(M01.Tables(0).Rows(0)("Qty"), "#.00")
                    Else
                        ' lblKP_Balance.Text = txtK_Qty.Text
                    End If

                Else
                    If UltraGrid4.Rows(i).Cells(7).Value = True Then
                        _fromDate = Month(UltraGrid4.Rows(i).Cells(1).Value) & "/" & Microsoft.VisualBasic.Day(UltraGrid4.Rows(i).Cells(1).Value) & "/" & Year(UltraGrid4.Rows(i).Cells(1).Value)
                        _Todate = Month(UltraGrid4.Rows(i).Cells(2).Value) & "/" & Microsoft.VisualBasic.Day(UltraGrid4.Rows(i).Cells(2).Value) & "/" & Year(UltraGrid4.Rows(i).Cells(2).Value)

                        ncQryType = "SKP"
                        vcWhere = ""
                        nvcFieldList1 = "(tmpRef_No," & "tmpMC_No," & "tmpGroup," & "tmpQuality," & "tmpSales_Order," & "tmpLine_Item," & "tmpKnt_Order," & "tmpBalance," & "tmpWeek_No," & "tmpYear," & "tmpStart_Time," & "tmpEnd_Time," & "tmpStatus," & "tmpUser," & "tmpStart_Date," & "tmpEnd_Date) " & "values(" & Delivary_Ref & ",'" & Trim(UltraGrid4.Rows(i).Cells(0).Value) & "','" & strMC_Group & "','" & txtQuality.Text & "','" & strSales_Order & "','" & strLine_Item & "','" & Trim(UltraGrid4.Rows(i).Cells(4).Value) & "','" & Trim(UltraGrid4.Rows(i).Cells(4).Value) & "'," & txtWeek_Knt.Text & ",'" & txtYear_Knt.Text & "','" & Trim(UltraGrid4.Rows(i).Cells(1).Value) & "','" & Trim(UltraGrid4.Rows(i).Cells(2).Value) & "','" & Trim(UltraGrid4.Rows(i).Cells(6).Value) & "','" & strDisname & "','" & _fromDate & "','" & _Todate & "')"
                        up_GetSetDelivary_Planning(ncQryType, nvcFieldList1, vcWhere, connection, transaction)


                        nvcFieldList1 = "select sum(tmpKnt_Order) as qty from tmpKnitting_Plan_Board where tmpRef_No=" & Delivary_Ref & "  and tmpSales_Order='" & strSales_Order & "' and tmpLine_Item='" & strLine_Item & "' and tmpWeek_No=" & txtWeek_Knt.Text & " and tmpYear=" & txtYear_Knt.Text & " group by tmpRef_No,tmpSales_Order,tmpLine_Item,tmpWeek_No,tmpYear"
                        M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                        If isValidDataset(M01) Then

                            lblKP_Balance.Text = CDbl(lblKP_Balance.Text) + M01.Tables(0).Rows(0)("Qty")
                            'lblKP_Balance.Text = Microsoft.VisualBasic.Format(M01.Tables(0).Rows(0)("Qty"), "#.00")
                        Else
                            ' lblKP_Balance.Text = txtK_Qty.Text
                        End If

                    End If

                End If
                i = i + 1
            Next
            transaction.Commit()


            nvcFieldList1 = "select sum(tmpKnt_Order) as Qty from tmpKnitting_Plan_Board where tmpSales_Order='" & strSales_Order & "' and tmpLine_Item=" & strLine_Item & " group by tmpSales_Order,tmpLine_Item"
            dsUser = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(dsUser) Then
                lblKP_Balance.Text = CDbl(lblKnt_Balance.Text) - CDbl(dsUser.Tables(0).Rows(0)("Qty"))
                txtK_Qty.Text = lblKP_Balance.Text
            End If
            connection.Close()
            'lblKP_Balance.Text = CDbl(lblKnt_Balance.Text) - lblKP_Balance.Text
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)

                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
                connection.Close()
            End If
        End Try
    End Function

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Call Save_Temp_KnittingBord()
    End Sub


    Private Sub txtEx_LibUse_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Call frmDelivaryQuatnew.CalculateBalance_To_Produce()
    End Sub

    Private Sub UltraButton29_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton29.Click
        If OPR_YDP.Visible = True Then
            OPR_YDP.Visible = False
        Else
            OPR_YDP.Visible = True
        End If
    End Sub

    Private Sub UltraButton30_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton30.Click
        If UltraGroupBox37.Visible = True Then
            UltraGroupBox37.Visible = False

        End If
    End Sub

    Private Sub UltraButton31_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton31.Click
        UltraGroupBox20.Visible = False
        With UltraGroupBox36
            .Width = 493
            .Height = 237
        End With
    End Sub

    Private Sub UltraButton32_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton32.Click
        Panel20.Visible = False
        chkShort.Checked = False
    End Sub

    Private Sub UltraButton33_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton33.Click
        Panel21.Visible = False
        chkShort.Checked = False
    End Sub

    Private Sub UltraButton34_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton34.Click
        Dim i As Integer
        Dim _McBooking_Status As Boolean

        'Insert Knitting Dash Board
        i = 0
        For Each uRow As UltraGridRow In UltraGrid4.Rows
            If UltraGrid4.Rows(i).Cells(7).Value = True Then
                _McBooking_Status = True
                'Calculate Week No

            End If
            i = i + 1
        Next
        '-------------------------------------------------------
        If _McBooking_Status = False Then
            MsgBox("Please allocate the machine", MsgBoxStyle.Information, "Information .....")

            Exit Sub
        End If
        Call Save_Temp_KnittingBord()
        Call Load_Gride_Knt()
    End Sub


    Private Sub dg_dye_Main_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dg_dye_Main.DoubleClick
        Dim _Rowindex As Integer
        'Dim _LineItem As String
        Dim vcWhere As String
        Dim M01 As DataSet
        Dim M02 As DataSet
        Dim M03 As DataSet
        Dim M04 As DataSet
        Dim i As Integer
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim Value As Double
        Dim _ST As String
        Dim Y As Integer
        Dim strWeek As String
        Dim Z As Integer
        Dim _BaseLine_Item As String
        Dim _BASEQUALITY As String
        Dim _Trim_Quality As String
        Dim _Qty As Integer

        Dim _caractRemove As String
        Dim _Rcode As String

        Try
            Call Load_Gride_Dye_Grige_Detailes()

            _Rowindex = dg_dye_Main.ActiveRow.Index
            _LineItem = Trim(dg_dye_Main.Rows(_Rowindex).Cells(1).Text)
            _base30class = Trim(dg_dye_Main.Rows(_Rowindex).Cells(2).Text)
            _Qty = 0
            _caractRemove = "-"
            _base30class = (Replace(_base30class, _caractRemove, ""))
            _DyeQuality = ""

            i = 0
            For Each uRow As UltraGridRow In dg_dye_Main.Rows
                If Trim(dg_dye_Main.Rows(i).Cells(0).Text) = True Then
                    If i = 0 Then
                        _LineItem = Trim(dg_dye_Main.Rows(i).Cells(1).Text)
                    Else
                        _LineItem = _LineItem & "','" & Trim(dg_dye_Main.Rows(i).Cells(1).Text)
                    End If
                End If
                i = i + 1
            Next

            vcWhere = "T01Sales_Order='" & strSales_Order & "' and T01Line_Item in ('" & _LineItem & "') and T01Status='C'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "T01D"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                _DyeQuality = M01.Tables(0).Rows(0)("M01Quality_No")
                If Trim(M01.Tables(0).Rows(0)("T01Maching")) <> "" Then
                    _BASEQUALITY = M01.Tables(0).Rows(0)("M01Quality_No")
                    _Qty = M01.Tables(0).Rows(0)("T01Qty")
                    _BaseLine_Item = M01.Tables(0).Rows(0)("T01Maching")
                    '_Trim_Quality = M01.Tables(0).Rows(0)("T01Maching")
                    '_DyeQuality = M01.Tables(0).Rows(0)("M01Quality_No")
                    vcWhere = "T01Sales_Order='" & strSales_Order & "' and T01Line_Item=" & Trim(M01.Tables(0).Rows(0)("T01Maching")) & " and T01Status='C'"
                    M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "T01D"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M02) Then
                        Value = 0
                        Dim newRow As DataRow = c_dataCustomer3_Dye.NewRow

                        newRow("##") = "Body"
                        newRow("Line Item") = M02.Tables(0).Rows(0)("T01Line_Item")
                        newRow("Material") = M02.Tables(0).Rows(0)("M01Material_No")
                        newRow("Description") = M02.Tables(0).Rows(0)("M01Quality")
                        _Trim_Quality = M02.Tables(0).Rows(0)("M01Quality_No")
                        _Qty = _Qty + M02.Tables(0).Rows(0)("T01Qty")
                        Value = _Qty
                        _ST = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        _ST = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                        newRow("Quantity") = _ST


                        Y = 0
                        vcWhere = "tmpSales_Order='" & strSales_Order & "' and tmpLine_Item=" & M02.Tables(0).Rows(0)("T01Line_Item") & ""
                        M03 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "SRK"), New SqlParameter("@vcWhereClause1", vcWhere))
                        For Each DTRow5 As DataRow In M03.Tables(0).Rows

                            vcWhere = "tmpSales_Order='" & strSales_Order & "' and tmpLine_Item=" & M02.Tables(0).Rows(0)("T01Line_Item") & " and tmpWeek_No=" & M03.Tables(0).Rows(Y)("tmpWeek_No") & ""
                            M04 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "SRK1"), New SqlParameter("@vcWhereClause1", vcWhere))
                            If isValidDataset(M04) Then
                                strWeek = "Week " & M03.Tables(0).Rows(Y)("tmpWeek_No")
                                Value = M04.Tables(0).Rows(0)("Qty")
                                _ST = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                                _ST = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                                newRow(strWeek) = _ST

                            End If
                            Y = Y + 1
                        Next

                        c_dataCustomer3_Dye.Rows.Add(newRow)
                    End If

                Else
                    i = 0
                    _BaseLine_Item = _LineItem
                    vcWhere = "T01Sales_Order='" & strSales_Order & "' and T01Line_Item in  ('" & _LineItem & "') and T01Status='C'"
                    M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "DPB"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M02) Then
                        'For Each DTRow3 As DataRow In M02.Tables(0).Rows
                        Value = 0
                        Dim newRow As DataRow = c_dataCustomer3_Dye.NewRow
                        _BASEQUALITY = M02.Tables(0).Rows(0)("M01Quality_No")
                        newRow("##") = "Body"
                        newRow("Line Item") = M02.Tables(0).Rows(0)("T01Line_Item")
                        newRow("Material") = M02.Tables(0).Rows(0)("M01Material_No")
                        newRow("Description") = M02.Tables(0).Rows(0)("M01Quality")
                        Value = M02.Tables(0).Rows(0)("T17Req_Griege")
                        _Qty = _Qty + M02.Tables(0).Rows(0)("T17Req_Griege")
                        _ST = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        _ST = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                        newRow("Quantity") = _ST


                        Y = 0
                        vcWhere = "tmpSales_Order='" & strSales_Order & "' and tmpLine_Item=" & Trim(M02.Tables(0).Rows(0)("T01Line_Item")) & ""
                        M03 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "SRK"), New SqlParameter("@vcWhereClause1", vcWhere))
                        For Each DTRow5 As DataRow In M03.Tables(0).Rows

                            vcWhere = "tmpSales_Order='" & strSales_Order & "' and tmpLine_Item=" & Trim(M02.Tables(0).Rows(0)("T01Line_Item")) & " and tmpWeek_No=" & M03.Tables(0).Rows(Y)("tmpWeek_No") & ""
                            M04 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "SRK1"), New SqlParameter("@vcWhereClause1", vcWhere))
                            If isValidDataset(M04) Then
                                strWeek = "Week " & M03.Tables(0).Rows(Y)("tmpWeek_No")
                                Value = M04.Tables(0).Rows(0)("Qty")
                                _ST = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                                _ST = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                                newRow(strWeek) = _ST

                            End If
                            Y = Y + 1
                        Next

                        c_dataCustomer3_Dye.Rows.Add(newRow)
                        '======================================================================
                        'Load Trim Quality
                        _Trim_Quality = ""
                        i = 0
                        vcWhere = "T01Sales_Order='" & strSales_Order & "' and T01Maching in ('" & _LineItem & "') and T01Status='C'"
                        M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "DPB"), New SqlParameter("@vcWhereClause1", vcWhere))
                        For Each DTRow6 As DataRow In M02.Tables(0).Rows
                            Value = 0
                            Dim newRow3 As DataRow = c_dataCustomer3_Dye.NewRow
                            'If i = 0 Then
                            '    _Trim_Quality = Trim(M02.Tables(0).Rows(i)("M01Quality_No"))
                            'Else
                            '    _Trim_Quality = _Trim_Quality & "','" & Trim(M02.Tables(0).Rows(i)("M01Quality_No"))
                            'End If
                            newRow3("##") = "Trim"
                            newRow3("Line Item") = M02.Tables(0).Rows(i)("T01Line_Item")
                            newRow3("Material") = M02.Tables(0).Rows(i)("M01Material_No")
                            newRow3("Description") = M02.Tables(0).Rows(i)("M01Quality")
                            Value = M02.Tables(0).Rows(i)("T17Req_Griege")
                            _Qty = _Qty + M02.Tables(0).Rows(i)("T17Req_Griege")
                            _ST = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                            _ST = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                            newRow3("Quantity") = _ST


                            Y = 0
                            vcWhere = "tmpSales_Order='" & strSales_Order & "' and tmpLine_Item=" & M02.Tables(0).Rows(0)("T01Line_Item") & ""
                            M03 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "SRK"), New SqlParameter("@vcWhereClause1", vcWhere))
                            For Each DTRow5 As DataRow In M03.Tables(0).Rows

                                vcWhere = "tmpSales_Order='" & strSales_Order & "' and tmpLine_Item=" & M02.Tables(0).Rows(0)("T01Line_Item") & " and tmpWeek_No=" & M03.Tables(0).Rows(Y)("tmpWeek_No") & ""
                                M04 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "SRK1"), New SqlParameter("@vcWhereClause1", vcWhere))
                                If isValidDataset(M04) Then
                                    strWeek = "Week " & M03.Tables(0).Rows(Y)("tmpWeek_No")
                                    Value = M04.Tables(0).Rows(0)("Qty")
                                    _ST = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                                    _ST = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                                    newRow3(strWeek) = _ST

                                End If
                                Y = Y + 1
                            Next

                            c_dataCustomer3_Dye.Rows.Add(newRow3)
                            i = i + 1
                        Next

                        Panel16.Visible = False
                        GroupBox1.Visible = False
                        dg_dye_Main.Visible = False
                        '---------------------------------------------------------------------------
                        'Loading Details NPL/PP/1stBulk
                        vcWhere = "T01Sales_Order='" & strSales_Order & "' and T01Line_Item in ('" & _BaseLine_Item & "') and T01Status='C'"
                        M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "T01D"), New SqlParameter("@vcWhereClause1", vcWhere))
                        If isValidDataset(M02) Then
                            txtDye_Bulk.Text = M02.Tables(0).Rows(0)("T01Bulk")
                            txtDye_LD.Text = M02.Tables(0).Rows(0)("T01Lab_Dye")
                            If Trim(M02.Tables(0).Rows(0)("T01Lab_Dye")) = "NOT APPROVED" Then
                                txtDye_LDApp.Text = M02.Tables(0).Rows(0)("T01POD")
                            End If
                            txtDye_NPL.Text = M02.Tables(0).Rows(0)("T01NPL")
                            If Trim(M02.Tables(0).Rows(0)("T01NPL")) = "NOT APPROVED" Then
                                txtDye_NPL_App.Text = M02.Tables(0).Rows(0)("T01NPL_AppDate")
                            End If

                            txtDye_PP.Text = M02.Tables(0).Rows(0)("T01PP")
                            If Trim(M02.Tables(0).Rows(0)("T01PP")) = "NOT APPROVED" Then
                                txtDye_PP_App.Text = M02.Tables(0).Rows(0)("T01PP_AppDate")
                            End If

                        End If

                    End If

                End If
            End If
            '====================================================================================
            'CHECK R-CODE

            vcWhere = "M16Material='" & _base30class & "'"
            M04 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "QRQ"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M04) Then
                _Rcode = Trim(M04.Tables(0).Rows(0)("M16R_Code"))
                vcWhere = "M14Order='" & Trim(M04.Tables(0).Rows(0)("M16R_Code")) & "'" ' and m14status='CUS APPD'"
                M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "RCOD"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M02) Then
                    'Using Trim Quality
                    vcWhere = "M49quality='" & _BASEQUALITY & "' and M49trim in ('" & _Trim_Quality & "')"
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "CGLI"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M01) Then
                        i = 0
                        lblConstrain.Text = M01.Tables(0).Rows(0)("M49Comment")
                        i = 0
                        For Each DTRow5 As DataRow In M01.Tables(0).Rows
                            lblConstrain.Text = M01.Tables(0).Rows(0)("M49Comment")
                            'MsgBox(Trim(M02.Tables(0).Rows(0)("M14grige")))

                            If Trim(M04.Tables(0).Rows(0)("M16Shade_Type")) = "White" Then
                                If Trim(M01.Tables(0).Rows(i)("M49Mc_Group")) = "LR" Then
                                    txtDye_S_LR.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49SR_White")))
                                    txtDye_D_LR.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49DR_White")))
                                ElseIf Trim(M01.Tables(0).Rows(i)("M49Mc_Group")) = "THAN" Then
                                    txtDye_S_Than.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49SR_White")))
                                    txtDye_D_Than.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49DR_White")))
                                ElseIf Trim(M01.Tables(0).Rows(i)("M49Mc_Group")) = "ECO" Then
                                    txtDye_S_Eco.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49SR_White")))
                                    txtDye_D_Eco.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49DR_White")))
                                ElseIf Trim(M01.Tables(0).Rows(i)("M49Mc_Group")) = "ECO +" Then
                                    txtS_Eco1.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49SR_White")))
                                    txtDye_D_Eco1.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49DR_White")))
                                End If
                            Else
                                If IsDBNull(M02.Tables(0).Rows(0)("M14Criticle")) Then
                                    If Trim(M01.Tables(0).Rows(i)("M49Mc_Group")) = "LR" Then
                                        txtDye_S_LR.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49SR_Nomal")))
                                        txtDye_D_LR.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49DR_Nomal")))
                                    ElseIf Trim(M01.Tables(0).Rows(i)("M49Mc_Group")) = "THAN" Then
                                        txtDye_S_Than.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49SR_Nomal")))
                                        txtDye_D_Than.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49DR_Nomal")))
                                    ElseIf Trim(M01.Tables(0).Rows(i)("M49Mc_Group")) = "ECO" Then
                                        txtDye_S_Eco.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49SR_Nomal")))
                                        txtDye_D_Eco.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49DR_Nomal")))
                                    ElseIf Trim(M01.Tables(0).Rows(i)("M49Mc_Group")) = "ECO +" Then
                                        txtS_Eco1.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49SR_Nomal")))
                                        txtDye_D_Eco1.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49DR_Nomal")))
                                    End If

                                ElseIf Trim(M02.Tables(0).Rows(0)("M14Criticle")) = "Y" Then
                                    If Trim(M01.Tables(0).Rows(i)("M49Mc_Group")) = "LR" Then
                                        txtDye_S_LR.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49SR_Critical")))
                                        txtDye_D_LR.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49DR_Critical")))
                                    ElseIf Trim(M01.Tables(0).Rows(i)("M49Mc_Group")) = "THAN" Then
                                        txtDye_S_Than.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49SR_Critical")))
                                        txtDye_D_Than.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49DR_Critical")))
                                    ElseIf Trim(M01.Tables(0).Rows(i)("M49Mc_Group")) = "ECO" Then
                                        txtDye_S_Eco.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49SR_Critical")))
                                        txtDye_D_Eco.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49DR_Critical")))
                                    ElseIf Trim(M01.Tables(0).Rows(i)("M49Mc_Group")) = "ECO +" Then
                                        txtS_Eco1.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49SR_Critical")))
                                        txtDye_D_Eco1.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49DR_Critical")))
                                    End If
                                ElseIf Trim(M02.Tables(0).Rows(0)("M14Criticle")) = "N" Or Trim(M02.Tables(0).Rows(0)("M14Criticle")) = "" Then
                                    If Trim(M01.Tables(0).Rows(i)("M49Mc_Group")) = "LR" Then
                                        txtDye_S_LR.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49SR_Nomal")))
                                        txtDye_D_LR.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49DR_Nomal")))
                                    ElseIf Trim(M01.Tables(0).Rows(i)("M49Mc_Group")) = "THAN" Then
                                        txtDye_S_Than.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49SR_Nomal")))
                                        txtDye_D_Than.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49DR_Nomal")))
                                    ElseIf Trim(M01.Tables(0).Rows(i)("M49Mc_Group")) = "ECO" Then
                                        txtDye_S_Eco.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49SR_Nomal")))
                                        txtDye_D_Eco.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49DR_Nomal")))
                                    ElseIf Trim(M01.Tables(0).Rows(i)("M49Mc_Group")) = "ECO +" Then
                                        txtS_Eco1.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49SR_Nomal")))
                                        txtDye_D_Eco1.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49DR_Nomal")))
                                    End If
                                End If
                            End If
                            ' MsgBox(Trim(M02.Tables(0).Rows(0)("M14Criticle")))
                            i = i + 1
                        Next

                        '_DyeQuality = _BASEQUALITY

                        'If Trim(_Trim_Quality) <> "" Then
                        '    _DyeQuality = _DyeQuality & "','" & _Trim_Quality
                        'Else

                        'End If

                        'Dye Quntity
                        'Grage Booking
                        'Developed by Suranga on 2016.7.15
                        _Qty = 0
                        vcWhere = "T12Sales_Order='" & strSales_Order & "' "
                        M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "T12Q"), New SqlParameter("@vcWhereClause1", vcWhere))
                        If isValidDataset(M01) Then
                            _Qty = M01.Tables(0).Rows(0)("Qty")
                        End If

                        'Knitting Booking
                        vcWhere = "tmpSales_Order='" & strSales_Order & "' "
                        M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "KPLB"), New SqlParameter("@vcWhereClause1", vcWhere))
                        If isValidDataset(M01) Then
                            _Qty = _Qty + M01.Tables(0).Rows(0)("Qty")
                        End If

                        lblDye_Qty.Text = (_Qty.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        lblDye_Qty.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", _Qty))

                        txtDye_S_LR.Appearance.BackColor = Color.White
                        txtDye_D_LR.Appearance.BackColor = Color.White
                        txtDye_S_Than.Appearance.BackColor = Color.White
                        txtDye_D_Than.Appearance.BackColor = Color.White
                        txtS_Eco1.Appearance.BackColor = Color.White
                        txtDye_D_Eco1.Appearance.BackColor = Color.White
                        txtDye_S_Eco.Appearance.BackColor = Color.White
                        txtDye_D_Eco.Appearance.BackColor = Color.White
                        lblPrevious_St_Code.Text = "-"
                        'PREVIOUS DYE MACHINE GROUP
                        vcWhere = "M16R_CODE='" & _Rcode & "' "
                        M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "DPBT"), New SqlParameter("@vcWhereClause1", vcWhere))
                        If isValidDataset(M01) Then
                            ' MsgBox(UCase(Trim(M01.Tables(0).Rows(0)("tmpGroup"))))
                            lblPrevious_St_Code.Text = Trim(M01.Tables(0).Rows(0)("tmpStock_Code"))
                            If UCase(Trim(M01.Tables(0).Rows(0)("tmpGroup"))) = "LR" Then
                                txtDye_S_LR.Appearance.BackColor = Color.Yellow
                                txtDye_D_LR.Appearance.BackColor = Color.Yellow
                            ElseIf UCase(Trim(M01.Tables(0).Rows(0)("tmpGroup"))) = "ECO" Then
                                txtDye_S_Eco.Appearance.BackColor = Color.Yellow
                                txtDye_D_Eco.Appearance.BackColor = Color.Yellow
                            ElseIf UCase(Trim(M01.Tables(0).Rows(0)("tmpGroup"))) = "THAN" Then
                                txtDye_S_Than.Appearance.BackColor = Color.Yellow
                                txtDye_D_Than.Appearance.BackColor = Color.Yellow
                            ElseIf UCase(Trim(M01.Tables(0).Rows(0)("tmpGroup"))) = "ECO +" Then
                                txtS_Eco1.Appearance.BackColor = Color.Yellow
                                txtDye_D_Eco1.Appearance.BackColor = Color.Yellow
                            End If
                            'End If
                        End If
                        DBEngin.CloseConnection(con)
                        con.ConnectionString = ""
                        con.close()

                        Call Load_Gride_Dye_Projection()
                        Call Data_Fill_Projection_Dye(_DyeQuality)

                    Else
                        'If Trim(M04.Tables(0).Rows(0)("m16shade_type")) = "Marls" Or Trim(M04.Tables(0).Rows(0)("m16shade_type")) = "Yarn Dyes" Then
                        '    txtDye_S_LR.Text = "200"
                        '    txtDye_S_Than.Text = "200"
                        '    txtDye_S_Eco.Text = "200"
                        '    txtDye_D_Eco.Text = "200"
                        '    txtDye_D_Eco1.Text = "200"
                        '    txtDye_D_LR.Text = "200"
                        '    txtDye_D_Than.Text = "200"
                        '    txtS_Eco1.Text = "200"

                        'End If
                        'Using Base Quality
                        vcWhere = "M49quality='" & _BASEQUALITY & "'"
                        M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "CGLI"), New SqlParameter("@vcWhereClause1", vcWhere))
                        i = 0
                        If isValidDataset(M01) Then
                            For Each DTRow5 As DataRow In M01.Tables(0).Rows
                                lblConstrain.Text = M01.Tables(0).Rows(0)("M49Comment")
                                'MsgBox(Trim(M02.Tables(0).Rows(0)("M14grige")))
                                If Trim(M02.Tables(0).Rows(0)("M14Criticle")) = "Y" And Trim(M02.Tables(0).Rows(0)("M14grige")) = "L" Then
                                    If Trim(M01.Tables(0).Rows(i)("M49Mc_Group")) = "LR" Then
                                        txtDye_S_LR.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49SR_Nomal")))
                                    ElseIf Trim(M01.Tables(0).Rows(i)("M49Mc_Group")) = "THAN" Then
                                        txtDye_S_Than.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49SR_Nomal")))
                                    ElseIf Trim(M01.Tables(0).Rows(i)("M49Mc_Group")) = "ECO" Then
                                        txtDye_S_Eco.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49SR_Nomal")))
                                    ElseIf Trim(M01.Tables(0).Rows(i)("M49Mc_Group")) = "ECO+" Then
                                        txtS_Eco1.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49SR_Nomal")))
                                    End If
                                ElseIf Trim(M02.Tables(0).Rows(0)("M14Criticle")) = "Y" And Trim(M02.Tables(0).Rows(0)("M14grige")) = "D" Then
                                    If Trim(M01.Tables(0).Rows(i)("M49Mc_Group")) = "LR" Then
                                        txtDye_S_LR.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49SR_Critical")))
                                    ElseIf Trim(M01.Tables(0).Rows(i)("M49Mc_Group")) = "THAN" Then
                                        txtDye_S_Than.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49SR_Critical")))
                                    ElseIf Trim(M01.Tables(0).Rows(i)("M49Mc_Group")) = "ECO" Then
                                        txtDye_S_Eco.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49SR_Critical")))
                                    ElseIf Trim(M01.Tables(0).Rows(i)("M49Mc_Group")) = "ECO+" Then
                                        txtS_Eco1.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49SR_Critical")))
                                    End If
                                ElseIf Trim(M02.Tables(0).Rows(0)("M14Criticle")) = "N" And Trim(M02.Tables(0).Rows(0)("M14grige")) = "D" Then
                                    If Trim(M01.Tables(0).Rows(i)("M49Mc_Group")) = "LR" Then
                                        txtDye_D_LR.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49DR_Critical")))
                                    ElseIf Trim(M01.Tables(0).Rows(i)("M49Mc_Group")) = "THAN" Then
                                        txtDye_D_Than.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49DR_Critical")))
                                    ElseIf Trim(M01.Tables(0).Rows(i)("M49Mc_Group")) = "ECO" Then
                                        txtDye_D_Eco.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49DR_Critical")))
                                    ElseIf Trim(M01.Tables(0).Rows(i)("M49Mc_Group")) = "ECO+" Then
                                        txtDye_D_Eco1.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49DR_Critical")))
                                    End If
                                ElseIf Trim(M02.Tables(0).Rows(0)("M14Criticle")) = "N" And Trim(M02.Tables(0).Rows(0)("M14grige")) = "L" Then
                                    If Trim(M01.Tables(0).Rows(i)("M49Mc_Group")) = "LR" Then
                                        txtDye_D_LR.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49DR_Nomal")))
                                    ElseIf Trim(M01.Tables(0).Rows(i)("M49Mc_Group")) = "THAN" Then
                                        txtDye_D_Than.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49DR_Nomal")))
                                    ElseIf Trim(M01.Tables(0).Rows(i)("M49Mc_Group")) = "ECO" Then
                                        txtDye_D_Eco.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49DR_Nomal")))
                                    ElseIf Trim(M01.Tables(0).Rows(i)("M49Mc_Group")) = "ECO+" Then
                                        txtDye_D_Eco1.Text = CInt(Trim(M01.Tables(0).Rows(i)("M49DR_Nomal")))
                                    End If
                                End If
                                i = i + 1
                            Next
                        Else
                            MsgBox("This Quality no not in capacity guide line(Dyeing).Temp taken 200kg per lane .....please cnt dyeing technical team", MsgBoxStyle.Information, "Information .........")
                            txtDye_S_LR.Text = "200"
                            txtDye_S_Than.Text = "200"
                            txtDye_S_Eco.Text = "200"
                            txtS_Eco1.Text = "200"
                            txtDye_S_LR.Text = "200"
                            txtDye_S_Than.Text = "200"
                            txtDye_S_Eco.Text = "200"
                            txtS_Eco1.Text = "200"
                            txtDye_D_LR.Text = "200"
                            txtDye_D_Than.Text = "200"
                            txtDye_D_Eco.Text = "200"
                            txtDye_D_Eco1.Text = "200"
                            txtDye_D_LR.Text = "200"
                            txtDye_D_Than.Text = "200"
                            txtDye_D_Eco.Text = "200"
                            txtDye_D_Eco1.Text = "200"
                        End If
                        '======================================================================================
                        'Projection 
                        _DyeQuality = _BASEQUALITY

                        'If _Trim_Quality <> "" Then

                        '    _DyeQuality = _DyeQuality & "','" & _Trim_Quality
                        'End If


                        lblDye_Qty.Text = (_Qty.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        lblDye_Qty.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", _Qty))

                        DBEngin.CloseConnection(con)
                        con.ConnectionString = ""
                        con.close()

                        Call Load_Gride_Dye_Projection()
                        Call Data_Fill_Projection_Dye(_DyeQuality)
                        End If
                Else
                        MsgBox("Can't Find the R-Code.Please Inform to the Merchant", MsgBoxStyle.Information, "Technova ......")
                        DBEngin.CloseConnection(con)
                        con.ConnectionString = ""
                        con.close()
                End If

            End If


        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)

                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
                con.close()
            End If
        End Try
    End Sub

    Private Sub UltraCheckEditor2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkNormal.CheckedChanged
        Panel20.Visible = True
    End Sub

   

    Private Sub chkShort_CheckedChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShort.CheckedChanged
        If chkShort.Checked = True Then
            Panel20.Visible = True
            ' Panel21.Visible = True
            Call Load_Dye_Mc()
        ElseIf chkNormal.Checked = False Then
            Panel20.Visible = False
            ' Panel21.Visible = False
        End If
    End Sub

 
    Private Sub chkN_Lead_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkN_Lead.CheckedChanged
        If chkN_Lead.Checked = True Then
            chkS_Lead.Checked = False
            OPRNLT.Visible = False
            Call Load_Grid_Delivary_Plan()
        End If
    End Sub

    Private Sub chkS_Lead_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkS_Lead.CheckedChanged
        If chkS_Lead.Checked = True Then
            chkN_Lead.Checked = False
            OPRNLT.Visible = True
            Call Load_Grid_Delivary_Plan_SLTime()
        End If
    End Sub

    

    Private Sub UltraButton35_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton35.Click
        Dim agroup5 As UltraGridGroup
        Dim _WeeNo As String
        Dim Value As Double
        Dim _STSting As String
        Dim _cOLUM As Integer

        Dim _Count As Integer
        If txtSL_Line.Text <> "" Then
        Else
            MsgBox("Please enter the Line Item", MsgBoxStyle.Information, "Information ......")
            Exit Sub
        End If

        If txtSL_Qty.Text <> "" Then
        Else
            MsgBox("Please enter the Qty", MsgBoxStyle.Information, "Information ......")
            Exit Sub
        End If

        If dgDel_Plan.DisplayLayout.Bands(0).Groups.Count = 3 Then

            agroup5 = dgDel_Plan.DisplayLayout.Bands(0).Groups.Add("DEL")
            agroup5.Header.Caption = "Delivery"
            'DYEING
            _cOLUM = Me.dgDel_Plan.DisplayLayout.Bands(0).Columns.Count
            '  _cOLUM = _cOLUM + 1

            _WeeNo = txtSL_Date.Text
            'Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(2).Width = 70 * _Count
            Me.dgDel_Plan.DisplayLayout.Bands(0).Columns.Add(_WeeNo, _WeeNo)
            Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(_WeeNo).Group = agroup5
            Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(_WeeNo).Width = 90
            Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(_WeeNo).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            Value = txtSL_Qty.Text
            _STSting = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _STSting = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

            dgDel_Plan.Rows(0).Cells(_cOLUM).Value = _STSting
            '_cOLUM = _cOLUM + 1
        Else
            _cOLUM = Me.dgDel_Plan.DisplayLayout.Bands(0).Columns.Count
            ' _cOLUM = _cOLUM + 1

            _WeeNo = txtSL_Date.Text
            'Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(2).Width = 70 * _Count
            Me.dgDel_Plan.DisplayLayout.Bands(0).Columns.Add(_WeeNo, _WeeNo)
            Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(_WeeNo).Group = agroup5
            Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(_WeeNo).Width = 90
            Me.dgDel_Plan.DisplayLayout.Bands(0).Columns(_WeeNo).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            Value = txtSL_Qty.Text
            _STSting = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _STSting = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

            dgDel_Plan.Rows(0).Cells(_cOLUM).Value = _STSting
            _cOLUM = _cOLUM + 1
        End If

    End Sub

    Private Sub UltraGrid4_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles UltraGrid4.InitializeLayout

    End Sub

    Function Check_Knt_MC() As Boolean
        Dim Sql As String
        Dim M01 As DataSet
        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Dim i As Integer
        Try
            i = UltraGrid4.ActiveRow.Index
            If UltraGrid4.Rows(i).Cells(7).Value = True Then
                Sql = "select * from tmpBlock_KntMC where tmpMC_No='" & UltraGrid4.Rows(i).Cells(0).Value & "'"
                M01 = DBEngin.ExecuteDataset(connection, transaction, Sql)
                If isValidDataset(M01) Then
                    If strDisname = Trim(M01.Tables(0).Rows(0)("tmpUser")) Then

                    Else
                        Check_Knt_MC = True
                        UltraGrid4.Rows(i).Cells(7).Value = False
                    End If
                Else
                    ' If Trim(M01.Tables(0).Rows(0)("tmpUser")) = strDisname Then
                    nvcFieldList1 = "Insert Into tmpBlock_KntMC(tmpRef_no,tmpSales_Order,tmpLine_Item,tmpMC_No,tmpUser,tmpTime)" & _
                                                           " values('" & Delivary_Ref & "', '" & strSales_Order & "','" & strLine_Item & "','" & UltraGrid4.Rows(i).Cells(0).Value & "','" & strDisname & "','" & Now & "')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                    'End If
                End If
            Else
                nvcFieldList1 = "delete from tmpBlock_KntMC where tmpMC_No='" & UltraGrid4.Rows(i).Cells(0).Value & "' and tmpUser='" & strDisname & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            End If

            transaction.Commit()
            connection.Close()
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.Close()
            End If
        End Try
    End Function

    Private Sub chkKnt_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkKnt.CheckedChanged
        If chkKnt.Checked = True Then
            txtWeek_Knt.ReadOnly = True
        Else
            txtWeek_Knt.ReadOnly = False
        End If
    End Sub



    Private Sub UltraButton37_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton37.Click
        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Dim nvcFiled As String
        Dim I As Integer
        Dim _DelivaryWeek As Integer
        Try
            nvcFieldList1 = "UPDATE tmpKnitting_Plan_Board SET tmpQ_Status='A' WHERE tmpSales_Order='" & strSales_Order & "' AND tmpLine_Item='" & _LineItem & "'"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            nvcFieldList1 = "select * from T18Delivary_Plane where "
            transaction.Commit()
            Panel16.Visible = True
            GroupBox1.Visible = True
            dg_dye_Main.Visible = True

            Panel17.Visible = False
            OPR_Dye.Visible = False


            Call Load_Gride_Dye_Capacity_New() 'Dye Capacity

            For I = 0 To 5
                Dim newRow3 As DataRow = c_dataCustomer4_Dye.NewRow
                If I = 0 Then
                    newRow3("##") = "Overroll Plant Filling(Kg)"
                ElseIf I = 1 Then
                    newRow3("##") = "Open Capacity(Kg)"
                ElseIf I = 2 Then
                    newRow3("##") = "Selected MC Group Filling(Kg)"
                ElseIf I = 3 Then
                    newRow3("##") = "Open Capacity Selected MC Group(Kg)"
                ElseIf I = 4 Then
                    newRow3("##") = "Flow Plan"
                ElseIf I = 5 Then
                    newRow3("##") = "Fix Plan"
                End If
                c_dataCustomer4_Dye.Rows.Add(newRow3)
            Next

         

            Call Load_Gride_Dye_Capacity()
            UltraTabControl1.Tabs(5).Enabled = True
            UltraTabControl1.SelectedTab = UltraTabControl1.Tabs(5)

            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
            connection.Close()
        Catch ex As EvaluateException
            If transactionCreated = False Then transaction.Rollback()
            MessageBox.Show(Me, ex.ToString)
            connection.Close()
        Finally
            If connectionCreated Then DBEngin.CloseConnection(connection)
        End Try
    End Sub

    Private Sub dg1_YB_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles dg1_YB.InitializeLayout

    End Sub
End Class