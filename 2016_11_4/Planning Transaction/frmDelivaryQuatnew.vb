Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports xl = Microsoft.Office.Interop.Excel
Imports System.Globalization
'Imports Office = Microsoft.Office.Core
Imports Microsoft.Office.Interop.Outlook
Imports System.Drawing
Imports Spire.XlS
Imports System.Xml
Imports System.IO
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet
Imports Infragistics.Win.UltraWinToolTip
Imports System.Globalization.CultureInfo
Public Class frmDelivaryQuatnew
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim c_dataCustomer1 As DataTable
    Dim _EPF As String
    Dim _Email As String
    Dim _LeadTime As String

    Dim c_dataCustomer As DataTable
    'Dim xlApp As New Excel.Application
    'Dim xlWBook As Excel.Workbook


    Private Sub frmDelivaryQuatnew_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        chkReq.Checked = True
        txtSpecial_Comment.ReadOnly = True
        Call Load_SalesOrder()
        Call Load_Gride_SalesOrder()

        Dim TipInfo2 As New UltraToolTipInfo()

        TipInfo2.ToolTipText = "Projection Over view"
        Me.UltraToolTipManager1.SetUltraToolTip(Me.UltraButton6, TipInfo2)
        Me.UltraToolTipManager1.DisplayStyle = Infragistics.Win.ToolTipDisplayStyle.BalloonTip
        txtPk_Dye.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtPk_Knt.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtYarn_Week.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
    End Sub

    Function MakeDataTable_Delivary_Quatation() As DataTable
        Dim I As Integer
        Dim X As Integer
        Dim _Lastweek As Integer


        ' MsgBox(DatePart("ww", Today))
        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Line Item", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Material", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Quality", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Qty", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Req Date", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Lead Time", GetType(String))
        dataTable.Columns.Add(colWork)
        ' colWork.MaxLength = 250
        'colWork = New DataColumn("P4P", GetType(Boolean))
        ''  colWork.MaxLength = 70
        'dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        'colWork = New DataColumn("Liability", GetType(Boolean))
        ''  colWork.MaxLength = 70
        'dataTable.Columns.Add(colWork)
        colWork = New DataColumn("Matching L/Items", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        Return dataTable
    End Function

    Function Load_Gride_SalesOrder()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer = MakeDataTable_Delivary_Quatation()
        UltraGrid1.DataSource = c_dataCustomer
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 50
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 80
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 210
            .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(3).Width = 60
            '.DisplayLayout.Bands(0).Columns(5).Width = 60
            '.DisplayLayout.Bands(0).Columns(8).Width = 60
            '.DisplayLayout.Bands(0).Columns(7).Width = 70
            '.DisplayLayout.Bands(0).Columns(9).Width = 60

        End With
    End Function

    Function Update_Transaction()
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

            ncQryType = "ADD"
            nvcFieldList1 = "(tmpRef_No," & "tmpSales_Order," & "tmpDate," & "tmpUser) " & "values(" & Delivary_Ref & ",'" & Trim(cboSO.Text) & "','" & Today & "','" & strDisname & "')"
            up_GetSetDelivary_Planning(ncQryType, nvcFieldList1, vcWhere, connection, transaction)

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

    Function Update_Records()
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

            nvcFieldList1 = "update P01PARAMETER set P01NO=P01NO +" & 1 & " where P01CODE='DP'"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

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

    Function Load_SalesOrder()
        Dim vcWhere As String
        Dim M01 As DataSet
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim T01 As DataSet
        Dim Value As Double
        Dim _REFNO As Integer

        Try
            vcWhere = "tmpSales_Order='" & Trim(cboSO.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                If strDisname = Trim(M01.Tables(0).Rows(0)("tmpUser")) Then
                    Delivary_Ref = M01.Tables(0).Rows(0)("tmpRef_No")
                Else
                    MsgBox("The Sales order use by " & Trim(M01.Tables(0).Rows(0)("tmpUser")), MsgBoxStyle.Information, "Information ....")
                    Exit Function
                End If
            Else
                If Trim(cboSO.Text) <> "" Then
                    Call Update_Records()
                    vcWhere = "P01CODE='DP'"
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "LST1"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M01) Then
                        Delivary_Ref = M01.Tables(0).Rows(0)("P01NO")
                    End If
                    Call Update_Transaction()
                End If
                End If





                If chkNon.Checked = True Then
                    vcWhere = "M01Status='A' "
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWhere))
                    With cboSO
                        .DataSource = M01
                        .Rows.Band.Columns(0).Width = 160
                        '.Rows.Band.Columns(1).Width = 260
                    End With

                    vcWhere = "M01Status='A' and M01Sales_Order='" & Trim(cboSO.Text) & "' "
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "LSTD"), New SqlParameter("@vcWhereClause1", vcWhere))
                    i = 0
                    For Each DTRow4 As DataRow In M01.Tables(0).Rows
                        Dim newRow As DataRow = c_dataCustomer.NewRow
                        newRow("Line Item") = M01.Tables(0).Rows(i)("M01Line_Item")
                        newRow("Material") = M01.Tables(0).Rows(i)("M01Material_No")
                        newRow("Quality") = M01.Tables(0).Rows(i)("M01Quality")
                        'newRow("Qty") = M01.Tables(0).Rows(i)("M01Delivary_Qty")
                        Value = M01.Tables(0).Rows(i)("M01SO_Qty")
                        Dim _Qty As String

                        _Qty = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        _Qty = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    newRow("Qty") = _Qty
                    newRow("Status") = False
                        '   newRow("Req Date") = Month(M01.Tables(0).Rows(i)("T01RQD")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T01RQD")) & "/" & Year(M01.Tables(0).Rows(i)("T01RQD"))
                        'newRow("P4P") = False
                        'newRow("Liability") = False


                        c_dataCustomer.Rows.Add(newRow)


                        i = i + 1
                    Next


                ElseIf chkReq.Checked = True Then
                    vcWhere = "T01Planner='" & strDisname & "' and T01Status='A' "
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "LST1"), New SqlParameter("@vcWhereClause1", vcWhere))
                    With cboSO
                        .DataSource = M01
                        .Rows.Band.Columns(0).Width = 160
                        '.Rows.Band.Columns(1).Width = 260
                    End With

                    vcWhere = "T01Planner='" & strDisname & "' and T01Status='A'  and T01Sales_Order='" & cboSO.Text & "'"
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "T01D"), New SqlParameter("@vcWhereClause1", vcWhere))
                    i = 0
                    For Each DTRow4 As DataRow In M01.Tables(0).Rows
                        Dim newRow As DataRow = c_dataCustomer.NewRow
                    newRow("Line Item") = M01.Tables(0).Rows(i)("T01Line_Item")
                    _REFNO = M01.Tables(0).Rows(i)("T01RefNo")
                        'Sql = "select * from M01Sales_Order_SAP where CONVERT(INT,M01Sales_Order)='" & Trim(cboSO.Text) & "' and M01Line_Item='" & Trim(M01.Tables(0).Rows(i)("T01Line_Item")) & "'"
                        'T01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    vcWhere = "M01Sales_Order='" & Trim(cboSO.Text) & "' and M01Line_Item=" & M01.Tables(0).Rows(i)("T01Line_Item") & ""
                        T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "LSTD"), New SqlParameter("@vcWhereClause1", vcWhere))
                        If isValidDataset(T01) Then
                            newRow("Material") = T01.Tables(0).Rows(0)("M01Material_No")
                            newRow("Quality") = T01.Tables(0).Rows(0)("M01Quality")

                        End If
                        Dim _Qty As String

                        Value = M01.Tables(0).Rows(i)("T01Qty")
                        _Qty = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        _Qty = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                        newRow("Qty") = _Qty
                        newRow("Req Date") = Month(M01.Tables(0).Rows(i)("T01RQD")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T01RQD")) & "/" & Year(M01.Tables(0).Rows(i)("T01RQD"))
                    newRow("Lead Time") = M01.Tables(0).Rows(i)("T01Lead_Time")
                    newRow("Matching L/Items") = M01.Tables(0).Rows(i)("T01Maching")
                    'newRow("P4P") = False
                        'newRow("Liability") = False


                        c_dataCustomer.Rows.Add(newRow)


                        i = i + 1
                    Next

                vcWhere = "T01_1Ref_No=" & _REFNO & ""
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "COMM"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    txtSpecial_Comment.Text = M01.Tables(0).Rows(0)("T01_1Comment")
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

    Private Sub chkReq_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkReq.CheckedChanged
        If chkReq.Checked = True Then
            chkNon.Checked = False
            cboSO.Text = ""
            Call Load_Gride_SalesOrder()
            Call Load_SalesOrder()

        End If
    End Sub

    Private Sub chkNon_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkNon.CheckedChanged
        If chkNon.Checked = True Then
            chkReq.Checked = False
            cboSO.Text = ""
            Call Load_Gride_SalesOrder()
            Call Load_SalesOrder()

        End If
    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        cboSO.Text = ""
        Call Load_Gride_SalesOrder()
    End Sub


    Private Sub cmdFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdFind.Click
        'Find Line Items for Perticuler Sales order
        Call Load_Gride_SalesOrder()
        Call Load_SalesOrder()


    End Sub

    Function Bulk_Knitting()
        Dim _Rowindex As Integer
        Dim _LineItem As Integer
        Dim i As Integer
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
        Try
            _Rowindex = UltraGrid1.ActiveRow.Index
            For Each uRow As UltraGridRow In UltraGrid1.Rows
                _LineItem = Trim(UltraGrid1.Rows(i).Cells(0).Text)
                If Trim(UltraGrid1.Rows(i).Cells(5).Text) = True Then
                    _LineItem = Trim(UltraGrid1.Rows(i).Cells(0).Text)
                    nvcFieldList1 = "select * from tmpBulk_Knitting_Plan where tmpRef_No=" & Delivary_Ref & " and tmpSales_Order='" & Trim(cboSO.Text) & "' and tmpLine_Item=" & _LineItem & ""
                    M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                    If isValidDataset(M01) Then
                    Else
                        nvcFieldList1 = "Insert Into tmpBulk_Knitting_Plan(tmpRef_No,tmpSales_Order,tmpLine_Item)" & _
                                                                  " values(" & Delivary_Ref & ", '" & Trim(cboSO.Text) & "'," & _LineItem & ")"
                        ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                    End If
                Else
                    nvcFieldList1 = "select * from tmpBulk_Knitting_Plan where tmpRef_No=" & Delivary_Ref & " and tmpSales_Order='" & Trim(cboSO.Text) & "' and tmpLine_Item=" & _LineItem & ""
                    M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                    If isValidDataset(M01) Then
                        nvcFieldList1 = "delete from tmpBulk_Knitting_Plan where tmpRef_No=" & Delivary_Ref & " and tmpSales_Order='" & Trim(cboSO.Text) & "' and tmpLine_Item=" & _LineItem & ""
                        M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                    Else

                    End If
                End If
                i = i + 1
            Next

            transaction.Commit()
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
    End Function

    Private Sub UltraGrid1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles UltraGrid1.Click
        ' Me.Hide()
        Dim _Rowindex As Integer
        Dim _LineItem As String
        Dim vcWhere As String
        Dim M01 As DataSet
        Dim M02 As DataSet

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim _30class As String

        Try
            If UltraGrid1.Rows.Count > 0 Then

            Else
                Exit Sub
            End If
            _Rowindex = UltraGrid1.ActiveRow.Index
            'If Trim(UltraGrid1.Rows(_Rowindex).Cells(5).Text) = True Then
            _LineItem = Trim(UltraGrid1.Rows(_Rowindex).Cells(0).Text)
            ' frmLoad_Pln.Show()
            strSales_Order = cboSO.Text
            strLine_Item = _LineItem
            frmKnitting_Plan_WithTab.MdiParent = MDIMain
            frmKnitting_Plan_WithTab.Show()
            With frmKnitting_Plan_WithTab
                .txtDate.Text = Today
                .txtDate.Visible = False
                .txtSO.Text = cboSO.Text
                strSales_Order = cboSO.Text
                .txtLine_Item.Text = _LineItem
                strLine_Item = _LineItem

                .txtQualityDis.Text = Trim(UltraGrid1.Rows(_Rowindex).Cells(2).Text)
                .txtMaterial.Text = Trim(UltraGrid1.Rows(_Rowindex).Cells(1).Text)
                str20Class = Trim(UltraGrid1.Rows(_Rowindex).Cells(1).Text)
                .txtOrder_Qty.Text = Trim(UltraGrid1.Rows(_Rowindex).Cells(3).Text)
                .txtFabric_Type.ReadOnly = True
                .txtConfact.ReadOnly = True
                .txtMaterial.ReadOnly = True
                .txtCLab.ReadOnly = True
                .txtWIP.ReadOnly = True
                .txtBalance.ReadOnly = True

                .txtSO.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
                .txtSO.ReadOnly = True
                .txtLine_Item.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
                .txtLine_Item.ReadOnly = True
                .txtOrder_Qty.ReadOnly = True
                .txtOrder_Qty.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
                .txtExcess_FG.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
                .txtWIP.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
                .txtBalance.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
                .txtQuality.ReadOnly = True
                .txtQualityDis.ReadOnly = True

                vcWhere = " M01Sales_Order='" & Trim(cboSO.Text) & "' and  M01Line_Item='" & Trim(UltraGrid1.Rows(_Rowindex).Cells(0).Text) & "'"
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "LSTD"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    .txtQuality.Text = M01.Tables(0).Rows(0)("M01Quality_No")
                    strQuality = M01.Tables(0).Rows(0)("M01Quality_No")
                End If

                'excess FG
                Dim _FG As Double
                Dim characterToRemove As String
                Dim _Balance As Double

                characterToRemove = "-"
                _30class = Trim(UltraGrid1.Rows(_Rowindex).Cells(1).Text)

                'MsgBox(Trim(fields(9)))
                _30class = (Replace(_30class, characterToRemove, ""))
                _FG = 0
                vcWhere = " m01sales_order<>'" & Trim(cboSO.Text) & "' and m0130class='" & _30class & "' and M01Status='I'"
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "FGS"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    _FG = M01.Tables(0).Rows(0)("M01Con")
                End If



                vcWhere = " M16Material='" & _30class & "'"
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "FSD"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    With frmKnitting_Plan_WithTab
                        .txtQuality.Text = M01.Tables(0).Rows(0)("M16Quality")
                        If Microsoft.VisualBasic.Left(M01.Tables(0).Rows(0)("M16Quality"), 1) = "Y" Then
                            .UltraTabControl1.Tabs(2).Enabled = True
                        Else
                            .UltraTabControl1.Tabs(2).Enabled = False
                        End If
                        strQuality = M01.Tables(0).Rows(0)("M16Quality")
                        '.txtMOQ.Text = (_Balance.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        '.txtMOQ.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", _Balance))
                    End With
                End If

                vcWhere = " m08location in ('2060','2055','2059','2062','2063','2065','2070') and m08meterial='" & _30class & "' and left(M08Sales_Order,2)<>'30'"
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "FGS1"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    _Balance = M01.Tables(0).Rows(0)("Qty")
                End If

                If _Balance > _FG Then
                    _FG = _Balance - _FG
                    .txtExcess_FG.Text = _FG
                    'Else
                    '    .txtExcess_FG.Text = "0"

                    _Balance = 0
                    vcWhere = "M09Oredr_Type in ('Exam','Finishing') and m09meterial='" & _30class & "' and left(M09Sales_Oredr,2)<>'30'"
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "ZPL"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M01) Then
                        _Balance = M01.Tables(0).Rows(0)("M09Qty_Mtr")
                    End If

                    vcWhere = "M09Oredr_Type='Dyeing' and Stock_Code<>'' and m09meterial='" & _30class & "' and left(M09Sales_Oredr,2)<>'30'"
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "ZPL1"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M01) Then
                        _Balance = _Balance + M01.Tables(0).Rows(0)("M09MTr")
                    End If

                    .txtWIP.Text = (_Balance.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    .txtWIP.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", _Balance))

                    '------------------------BALANCE QTY
                    _Balance = 0
                    If IsNumeric(.txtUse_FG.Text) Then
                        _Balance = CDbl(.txtUse_FG.Text)
                    End If

                    If IsNumeric(.txtUse_WIP.Text) Then
                        _Balance = _Balance + CDbl(.txtUse_WIP.Text)
                    End If

                    _Balance = CDbl(.txtOrder_Qty.Text) - _Balance
                    .txtBalance.Text = (_Balance.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    .txtBalance.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", _Balance))
                Else
                    Dim _EXFg As Double
                    _EXFg = 0
                    'CHANGE REQUIRMENT USING AMILA
                    _EXFg = _Balance - _FG

                    _Balance = 0
                    vcWhere = "M09Oredr_Type in ('Exam','Finishing') and m09meterial='" & _30class & "' and left(M09Sales_Oredr,2)<>'30'"
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "ZPL"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M01) Then
                        _Balance = M01.Tables(0).Rows(0)("M09Qty_Mtr")
                    End If

                    vcWhere = "M09Oredr_Type='Dyeing' and Stock_Code<>'' and m09meterial='" & _30class & "' and left(M09Sales_Oredr,2)<>'30'"
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "ZPL1"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M01) Then
                        _Balance = _Balance + M01.Tables(0).Rows(0)("M09MTr")
                    End If

                    '.txtWIP.Text = (_Balance.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    '.txtWIP.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", _Balance))

                    If _FG < (_Balance + _EXFg) Then
                        .txtExcess_FG.Text = "0"
                        _Balance = (CDbl(.txtWIP.Text) + _EXFg) - _FG
                        .txtWIP.Text = (_Balance.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        .txtWIP.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", _Balance))
                    End If

                    _Balance = 0
                    If IsNumeric(.txtUse_FG.Text) Then
                        _Balance = CDbl(.txtUse_FG.Text)
                    End If

                    If IsNumeric(.txtUse_WIP.Text) Then
                        _Balance = _Balance + CDbl(.txtUse_WIP.Text)
                    End If

                    _Balance = CDbl(.txtOrder_Qty.Text) - _Balance
                    .txtBalance.Text = (_Balance.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    .txtBalance.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", _Balance))

                End If

                'CRITICLE
                vcWhere = " M16Material='" & _30class & "'"
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "RCODE"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    vcWhere = "M14Order='" & M01.Tables(0).Rows(0)("M16R_Code") & "'"
                    M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "RCDE1"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M02) Then

                        If IsDBNull(M02.Tables(0).Rows(0)("M14Criticle")) Then
                            With frmKnitting_Plan_WithTab
                                .chkCry2.Checked = True

                            End With
                        Else
                            If Trim(M02.Tables(0).Rows(0)("M14Criticle")) = "Y" Then
                                With frmKnitting_Plan_WithTab
                                    .chkCry1.Checked = True

                                End With
                            Else
                                With frmKnitting_Plan_WithTab
                                    .chkCry2.Checked = True

                                End With
                            End If
                        End If
                    End If
                End If
                '-------------------------------------------------------------
                'lab dip/NPL/PP
                vcWhere = " T01Sales_Order=" & cboSO.Text & " and T01Line_Item='" & _LineItem & "'"
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "LAB"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    With frmKnitting_Plan_WithTab
                        '.txtFabric_Type.Text = Trim(M01.Tables(0).Rows(0)("T01Bulk"))
                        .lblDate.Text = Trim(M01.Tables(0).Rows(0)("T01RQD"))
                        .txt1stBulk.Text = Trim(M01.Tables(0).Rows(0)("T01Bulk"))
                        .lblBulk_Dis.Text = Trim(M01.Tables(0).Rows(0)("T01Bulk")) & " Order"
                        .lblBuld_Dis1.Text = Trim(M01.Tables(0).Rows(0)("T01Bulk")) & " Order"

                        .txtPk_Biz.Text = Trim(M01.Tables(0).Rows(0)("T01Pack_Biz"))
                        .txtPk_Biz1.Text = Trim(M01.Tables(0).Rows(0)("T01Pack_Biz"))
                        .txtPk_Biz2.Text = Trim(M01.Tables(0).Rows(0)("T01Pack_Biz"))
                    End With
                    If Trim(M01.Tables(0).Rows(0)("T01Lab_Dye")) = "APPROVED" Then
                        With frmKnitting_Plan_WithTab
                            .chkLab1.Checked = True

                        End With
                    Else
                        With frmKnitting_Plan_WithTab
                            .chkLab2.Checked = True
                            .txtDate.Visible = True
                            .txtDate.Text = Trim(M01.Tables(0).Rows(0)("T01POD"))
                        End With
                    End If
                    '------------------------------------------------------------
                    If Trim(M01.Tables(0).Rows(0)("T01NPL")) = "APPROVED" Then
                        With frmKnitting_Plan_WithTab
                            .chkNPL1.Checked = True
                            .txtK_NPL.Text = "APPROVED"
                        End With
                    Else
                        With frmKnitting_Plan_WithTab
                            .chkNPL2.Checked = True
                            .txtK_NPL.Text = "NO"
                        End With
                    End If
                    '--------------------------------------------------
                    If Trim(M01.Tables(0).Rows(0)("T01PP")) = "APPROVED" Then
                        With frmKnitting_Plan_WithTab
                            .chkPP1.Checked = True

                        End With
                    Else
                        With frmKnitting_Plan_WithTab
                            .chkPP2.Checked = True

                        End With
                    End If
                End If
                '----------------------------------------------------------
                'DYE WASTAGE
                vcWhere = " M24Material='" & _30class & "'"
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "BOM"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    With frmKnitting_Plan_WithTab
                        .txtDye_Wast.Text = M01.Tables(0).Rows(0)("M24WST") * 100
                    End With
                End If
                '----------------------------------------------------------
                'Color Lab
                vcWhere = " M16Material='" & _30class & "'"
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "CLAB"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    With frmKnitting_Plan_WithTab
                        .txtCLab.Text = M01.Tables(0).Rows(0)("M14Status")

                    End With
                End If
                'RCODE
                vcWhere = " M16Material='" & _30class & "'"
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "RCD"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    With frmKnitting_Plan_WithTab
                        .txtRcode.Text = M01.Tables(0).Rows(0)("M16R_Code")
                        .txtFabric_Type.Text = Trim(M01.Tables(0).Rows(0)("M16Product_Type"))
                        .txtShade.Text = Trim(M01.Tables(0).Rows(0)("M16Product_Type"))
                    End With
                End If
                '----------------------------------------------------------
                'Convention Factor
                vcWhere = " M22Quality='" & .txtQuality.Text & "'"
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "CON"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    With frmKnitting_Plan_WithTab
                        .txtConfact.Text = M01.Tables(0).Rows(0)("M22Con_Fact")
                        .txtFabrication.Text = M01.Tables(0).Rows(0)("M22Fabric_Type")
                    End With
                End If

                '----------------------------------------------------------------------
                'MOQ
                vcWhere = " M31Quality='" & .txtQuality.Text & "'"
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "MOQ"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    With frmKnitting_Plan_WithTab
                        _Balance = M01.Tables(0).Rows(0)("M31Qty")
                        .txtMOQ.Text = (_Balance.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        .txtMOQ.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", _Balance))
                    End With
                End If
                '-----------------------------------------------------------------------
                'Fabric Shade
                Dim _Material As String
                _Material = .txtMaterial.Text
                characterToRemove = "-"

                'MsgBox(Trim(fields(9)))
                _Material = (Replace(_Material, characterToRemove, ""))


                vcWhere = " M16Material='" & _Material & "'"
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "FSD"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    With frmKnitting_Plan_WithTab
                        .txtFabric_Shade.Text = M01.Tables(0).Rows(0)("M16Shade_Type")
                        '.txtMOQ.Text = (_Balance.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        '.txtMOQ.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", _Balance))
                    End With
                End If

                'LIABILITY
                If IsNumeric(frmKnitting_Plan_WithTab.txtMOQ.Text) Then
                    With frmKnitting_Plan_WithTab
                        If CDbl(.txtMOQ.Text) >= CDbl(.txtOrder_Qty.Text) Then
                            _Balance = CDbl(.txtMOQ.Text) - CDbl(.txtOrder_Qty.Text)
                            .txtLIB.Text = (_Balance.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                            .txtLIB.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", _Balance))
                        Else
                            .txtLIB.Text = "0"
                        End If
                    End With
                End If
                'Requied Grige Qty
                If IsNumeric(frmKnitting_Plan_WithTab.txtBalance.Text) Then
                    With frmKnitting_Plan_WithTab
                        If IsNumeric(.txtConfact.Text) And IsNumeric(.txtDye_Wast.Text) Then
                            _Balance = CDbl(.txtBalance.Text) / CDbl(.txtConfact.Text)
                            _Balance = _Balance * CDbl(.txtDye_Wast.Text)
                            _Balance = _Balance / 100
                            _Balance = _Balance + CDbl(.txtBalance.Text)
                            .txtReq_Grg.Text = (_Balance.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                            .txtReq_Grg.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", _Balance))
                        End If
                    End With
                End If

                '----------------------------------------------------
                'Requierd Liability
                If IsNumeric(frmKnitting_Plan_WithTab.txtLIB.Text) Then
                    With frmKnitting_Plan_WithTab
                        If IsNumeric(.txtConfact.Text) And IsNumeric(.txtDye_Wast.Text) Then
                            _Balance = CDbl(.txtLIB.Text) / CDbl(.txtConfact.Text)
                            _Balance = _Balance * CDbl(.txtDye_Wast.Text)
                            _Balance = _Balance / 100
                            _Balance = _Balance + (CDbl(.txtLIB.Text) / CDbl(.txtConfact.Text))
                            .txtReg_LIb.Text = (_Balance.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                            .txtReg_LIb.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", _Balance))
                        End If
                    End With
                End If

                '-------------------------------------------------------------------------
                'Exsist LIB
                Dim Value As Double

                vcWhere = " m08Meterial='" & _Material & "' and m08Location in ('2060','2059','2065','2070','2062','2050','6060') and left(M08Sales_Order,2)='30'"
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "ELI"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    With frmKnitting_Plan_WithTab
                        Value = M01.Tables(0).Rows(0)("m08Qty_mtr")
                        .txtEx_Lib.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        .txtEx_Lib.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                        .txtEx_LibUse.Text = .txtEx_Lib.Text
                        '.txtMOQ.Text = (_Balance.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        '.txtMOQ.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", _Balance))
                    End With
                End If

                vcWhere = " m09Meterial='" & _Material & "' and M09Oredr_Type in ('Finishing','Exam') and left(M09Sales_Oredr,2)='30'"
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "ELE"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    With frmKnitting_Plan_WithTab
                        Value = CDbl(.txtEx_Lib.Text) + M01.Tables(0).Rows(0)("m09Qty_Mtr")
                        .txtEx_Lib.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        .txtEx_Lib.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                        .txtEx_LibUse.Text = .txtEx_Lib.Text
                        '.txtMOQ.Text = (_Balance.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        '.txtMOQ.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", _Balance))
                    End With
                End If

                vcWhere = " m09Meterial='" & _Material & "' and M09Oredr_Type in ('Dyeing') and left(M09Sales_Oredr,2)='30'"
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "ELD"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    With frmKnitting_Plan_WithTab
                        Value = CDbl(.txtEx_Lib.Text) + M01.Tables(0).Rows(0)("m09Qty_Mtr")
                        .txtEx_Lib.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        .txtEx_Lib.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                        .txtEx_LibUse.Text = .txtEx_Lib.Text
                        '.txtMOQ.Text = (_Balance.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        '.txtMOQ.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", _Balance))
                    End With
                End If
                .txtUse_FG.Text = "0"
            End With

            Call Search_DY_KNTWeek(cboSO.Text)
            Call CalculateBalance_To_Produce()
            '  Call Bulk_Knitting()
            Call frmKnitting_Plan_WithTab.Load_Projection()
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
    End Sub

    Function CalculateBalance_To_Produce()
        On Error Resume Next
        Dim _Balance As Double

        _Balance = 0
        With frmKnitting_Plan_WithTab
            If IsNumeric(.txtUse_FG.Text) Then
                _Balance = CDbl(.txtUse_FG.Text)
            End If

            If IsNumeric(.txtUse_WIP.Text) Then
                _Balance = _Balance + CDbl(.txtUse_WIP.Text)
            End If

            _Balance = CDbl(.txtOrder_Qty.Text) - _Balance
            _Balance = _Balance - CDbl(.txtEx_LibUse.Text)
            .txtBalance.Text = (_Balance.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            .txtBalance.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", _Balance))
        End With

        'Requied Grige Qty
        If IsNumeric(frmKnitting_Plan_WithTab.txtBalance.Text) Then
            With frmKnitting_Plan_WithTab
                If .txtDye_Wast.Text <> "" Then
                Else
                    .txtDye_Wast.Text = "0"
                End If
                If IsNumeric(.txtConfact.Text) And IsNumeric(.txtDye_Wast.Text) Then
                    _Balance = CDbl(.txtBalance.Text) / CDbl(.txtConfact.Text)
                    _Balance = _Balance * 100
                    _Balance = _Balance / (100 - CDbl(.txtDye_Wast.Text))
                    ' _Balance = _Balance + CDbl(.txtBalance.Text)
                    .txtReq_Grg.Text = (_Balance.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    .txtReq_Grg.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", _Balance))
                End If
            End With
        End If

        '----------------------------------------------------
        'Reueried Liability

        If IsNumeric(frmKnitting_Plan_WithTab.txtLIB.Text) Or frmKnitting_Plan_WithTab.txtLIB.Text = "" Then
            With frmKnitting_Plan_WithTab
                If IsNumeric(.txtConfact.Text) And IsNumeric(.txtDye_Wast.Text) Then
                    If (.txtLIB.Text) <> "" Then
                        _Balance = CDbl(.txtLIB.Text) / CDbl(.txtConfact.Text)
                        _Balance = _Balance * CDbl(.txtDye_Wast.Text)
                        _Balance = _Balance / 100
                        _Balance = _Balance + (CDbl(.txtLIB.Text) / CDbl(.txtConfact.Text))
                        .txtReg_LIb.Text = (_Balance.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        .txtReg_LIb.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", _Balance))
                    End If
                End If
            End With
        End If

    End Function

    Private Sub UltraButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton6.Click
        If OPR11.Visible = True Then
            OPR11.Visible = False
            UltraGroupBox4.Visible = False
        Else
            OPR11.Visible = True
            UltraGroupBox4.Visible = True
            Call Load_Gride2()
        End If
    End Sub

    Function Load_Gride2()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim T01 As DataSet
        Dim T02 As DataSet
        Dim X As Integer
        Dim _Code As Integer

        Dim vcWhere As String
        Dim i As Integer
        Dim agroup1 As UltraGridGroup
        Dim agroup2 As UltraGridGroup
        Dim agroup3 As UltraGridGroup
        Dim agroup4 As UltraGridGroup
        Dim agroup5 As UltraGridGroup
        Dim agroup6 As UltraGridGroup

        Dim _rowcount As Integer
        Dim _delivaryDate As Date
        Dim _Coloum_Count As Integer
        Dim _Row_Count As Integer
        Dim _projection_Qty As Double
        Dim _1stBulk_Status As Boolean

        Try
            UltraGrid2.DisplayLayout.Bands(0).Groups.Clear()
            UltraGrid2.DisplayLayout.Bands(0).Columns.Dispose()


            agroup1 = UltraGrid2.DisplayLayout.Bands(0).Groups.Add("GroupH")
            agroup1.Header.Caption = "Quality"

            agroup1.Width = 70



            Dim dt As DataTable = New DataTable()
            ' dt.Columns.Add("ID", GetType(Integer))
            Dim colWork As New DataColumn("##", GetType(String))
            dt.Columns.Add(colWork)
            colWork.ReadOnly = True


            'Load Quality
            vcWhere = "T15Sales_Order='" & Trim(cboSO.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "PROV"), New SqlParameter("@vcWhereClause1", vcWhere))
            i = 0
            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                dt.Rows.Add(Trim(M01.Tables(0).Rows(i)("T15Code")) & "-" & Trim(M01.Tables(0).Rows(i)("T15Shade")))
                dt.Rows.Add(Trim(M01.Tables(0).Rows(i)("T15Code")) & "-" & Trim(M01.Tables(0).Rows(i)("T15Shade")))
                dt.Rows.Add(Trim(M01.Tables(0).Rows(i)("T15Code")) & "-" & Trim(M01.Tables(0).Rows(i)("T15Shade")))
                dt.Rows.Add(Trim(M01.Tables(0).Rows(i)("T15Code")) & "-" & Trim(M01.Tables(0).Rows(i)("T15Shade")))

                i = i + 1
            Next


            Me.UltraGrid2.SetDataBinding(dt, Nothing)
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns(0).Group = agroup1
            ' Me.UltraGrid2.DisplayLayout.Bands(0).Columns(1).Group = agroup1
            'Me.dg_Knt_Week.DisplayLayout.Bands(0).Columns(1).Group = agroup1
            'Me.dg_Knt_Week.DisplayLayout.Bands(0).Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns(0).Width = 180
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

            'Me.UltraGrid2.SetDataBinding(dt1, Nothing)
            'Me.UltraGrid2.DisplayLayout.Bands(0).Columns(0).Group = agroup2
            Me.UltraGrid2.DisplayLayout.Override.MergedCellStyle = MergedCellStyle.Always
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns(0).MergedCellEvaluationType = MergedCellEvaluationType.MergeSameText


            agroup2 = UltraGrid2.DisplayLayout.Bands(0).Groups.Add("Group1")
            agroup2.Header.Caption = "##"
            agroup2.Width = 70
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns.Add("Line", "")
            Me.UltraGrid2.DisplayLayout.Bands(0).Columns("Line").Group = agroup2

            i = 0
            _rowcount = 0
            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                'If _rowcount = 0 Then
                UltraGrid2.Rows(_rowcount).Cells(1).Value = "Projection"
                _rowcount = _rowcount + 1
                UltraGrid2.Rows(_rowcount).Cells(1).Value = "Filling"
                ' UltraGrid2.Rows(_rowcount).Cells(0).Text = ""
                _rowcount = _rowcount + 1
                UltraGrid2.Rows(_rowcount).Cells(1).Value = "Open"
                'UltraGrid2.Rows(_rowcount).Cells(0).Value = ""
                UltraGrid2.Rows(_rowcount + 1).Cells(1).Appearance.BackColor = Color.Gold
                UltraGrid2.Rows(_rowcount + 1).Cells(0).Appearance.BackColor = Color.Gold
                _rowcount = _rowcount + 2
                'End If
                'UltraGrid2.Rows(_rowcount).Cells(1).Appearance.BackColor = Color.ForestGreen
                'UltraGrid2.Rows(_rowcount).Cells(0).Appearance.BackColor = Color.ForestGreen
                i = i + 1
            Next
            Dim _WeekNo As Integer
            _WeekNo = 0
            i = 0
            'YARN DYE
            vcWhere = "T15Sales_Order='" & Trim(cboSO.Text) & "' AND LEFT(M01Quality_No,1)='Y'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "PROV"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                vcWhere = "T15Sales_Order='" & Trim(cboSO.Text) & "' AND LEFT(M01Quality_No,1)='Y' "
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "PROX"), New SqlParameter("@vcWhereClause1", vcWhere))
                For Each DTRow3 As DataRow In M01.Tables(0).Rows
                    Dim remain As Integer
                    Dim noOfWeek As Integer
                    Dim userDate As Date
                    Dim _LastDate As Date
                    Dim _TimeSpan As TimeSpan

                    userDate = DateTime.Parse(M01.Tables(0).Rows(i)("T15Month") & "/1/" & M01.Tables(0).Rows(i)("T15Year"))
                    ' MsgBox(WeekdayName(Weekday(userDate)))
                    If WeekdayName(Weekday(userDate)) = "Sunday" Then
                        userDate = userDate.AddDays(-4)
                    ElseIf WeekdayName(Weekday(userDate)) = "Monday" Then
                        userDate = userDate.AddDays(-5)
                    ElseIf WeekdayName(Weekday(userDate)) = "Tuesday" Then
                        userDate = userDate.AddDays(-6)
                    ElseIf WeekdayName(Weekday(userDate)) = "Thusday" Then
                        userDate = userDate.AddDays(-1)
                    ElseIf WeekdayName(Weekday(userDate)) = "Friday" Then
                        userDate = userDate.AddDays(-2)
                    ElseIf WeekdayName(Weekday(userDate)) = "Saturday" Then
                        userDate = userDate.AddDays(-3)

                    End If


                    _LastDate = DateTime.Parse(M01.Tables(0).Rows(i)("T15Month") & "/1/" & M01.Tables(0).Rows(i)("T15Year"))
                    ' MsgBox(Date.DaysInMonth(_LastDate.Year, _LastDate.Month))
                    _LastDate = DateTime.Parse(M01.Tables(0).Rows(i)("T15Month") & "/" & Date.DaysInMonth(_LastDate.Year, _LastDate.Month) & "/" & M01.Tables(0).Rows(i)("T15Year"))
                    _TimeSpan = _LastDate.Subtract(userDate)
                    _WeekNo = _WeekNo + (_TimeSpan.Days / 7)
                    i = i + 1
                Next

                agroup6 = UltraGrid2.DisplayLayout.Bands(0).Groups.Add("Group6")
                agroup6.Header.Caption = "Yarn Dye"
                _WeekNo = _WeekNo * 60
                agroup6.Width = _WeekNo

            End If

            i = 0
            vcWhere = "T15Sales_Order='" & Trim(cboSO.Text) & "' AND LEFT(M01Quality_No,1)='Y' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "PROX"), New SqlParameter("@vcWhereClause1", vcWhere))
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
                    ' userDate = userDate.AddDays(-1)
                ElseIf WeekdayName(Weekday(userDate)) = "Friday" Then
                    userDate = userDate.AddDays(-1)
                ElseIf WeekdayName(Weekday(userDate)) = "Saturday" Then
                    userDate = userDate.AddDays(-2)

                End If
                _LastDate = DateTime.Parse(M01.Tables(0).Rows(i)("T15Month") & "/1/" & M01.Tables(0).Rows(i)("T15Year"))
                ' MsgBox(Date.DaysInMonth(_LastDate.Year, _LastDate.Month))
                _LastDate = DateTime.Parse(M01.Tables(0).Rows(i)("T15Month") & "/" & Date.DaysInMonth(_LastDate.Year, _LastDate.Month) & "/" & M01.Tables(0).Rows(i)("T15Year"))
                _TimeSpan = _LastDate.Subtract(userDate)
                _WeekNo = _TimeSpan.Days / 7

                userDate = userDate.AddDays(+7)
                vcWhere = "T15Sales_Order='" & Trim(cboSO.Text) & "' AND t01bulk='1st Bulk' and T15Year=" & M01.Tables(0).Rows(i)("T15Year") & " and T15Month=" & M01.Tables(0).Rows(i)("T15Month") & ""
                T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "FSTB"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(T01) Then
                    userDate = userDate.AddDays(-35)
                    _1stBulk_Status = True
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

                    _StrWeek = "Week- " & intWeek
                    _StrWeek1 = "Week " & intWeek
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns.Add(_StrWeek, _StrWeek1)
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).Group = agroup6
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).Width = 60
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                    userDate = userDate.AddDays(+7)
                    culture = System.Globalization.CultureInfo.CurrentCulture
                    intWeek = culture.Calendar.GetWeekOfYear(userDate, System.Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday)

                    _StrWeek = "Week- " & intWeek
                    _StrWeek1 = "Week " & intWeek
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns.Add(_StrWeek, _StrWeek1)
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).Group = agroup6
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).Width = 60
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                    userDate = userDate.AddDays(+7)
                    culture = System.Globalization.CultureInfo.CurrentCulture
                    intWeek = culture.Calendar.GetWeekOfYear(userDate, System.Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday)

                    _StrWeek = "Week- " & intWeek
                    _StrWeek1 = "Week " & intWeek
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns.Add(_StrWeek, _StrWeek1)
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).Group = agroup6
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).Width = 60
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                    userDate = userDate.AddDays(+7)
                    culture = System.Globalization.CultureInfo.CurrentCulture
                    intWeek = culture.Calendar.GetWeekOfYear(userDate, System.Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday)

                    _StrWeek = "Week- " & intWeek
                    _StrWeek1 = "Week " & intWeek
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns.Add(_StrWeek, _StrWeek1)
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).Group = agroup6
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).Width = 60
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                    userDate = userDate.AddDays(+7)
                    culture = System.Globalization.CultureInfo.CurrentCulture
                    intWeek = culture.Calendar.GetWeekOfYear(userDate, System.Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday)

                    _StrWeek = "Week- " & intWeek
                    _StrWeek1 = "Week " & intWeek
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns.Add(_StrWeek, _StrWeek1)
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).Group = agroup6
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).Width = 60
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                ElseIf _WeekNo = 4 Then
                    Dim culture As System.Globalization.CultureInfo
                    Dim intWeek As Integer
                    Dim _StrWeek As String
                    Dim _StrWeek1 As String
                    culture = System.Globalization.CultureInfo.CurrentCulture
                    intWeek = culture.Calendar.GetWeekOfYear(userDate, System.Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday)

                    _StrWeek = "Week- " & intWeek
                    _StrWeek1 = "Week " & intWeek
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns.Add(_StrWeek, _StrWeek1)
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).Group = agroup6
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).Width = 60
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                    userDate = userDate.AddDays(+7)
                    culture = System.Globalization.CultureInfo.CurrentCulture
                    intWeek = culture.Calendar.GetWeekOfYear(userDate, System.Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday)

                    _StrWeek = "Week- " & intWeek
                    _StrWeek1 = "Week " & intWeek
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns.Add(_StrWeek, _StrWeek1)
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).Group = agroup6
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).Width = 60
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                    userDate = userDate.AddDays(+7)
                    culture = System.Globalization.CultureInfo.CurrentCulture
                    intWeek = culture.Calendar.GetWeekOfYear(userDate, System.Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday)

                    _StrWeek = "Week- " & intWeek
                    _StrWeek1 = "Week " & intWeek
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns.Add(_StrWeek, _StrWeek1)
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).Group = agroup6
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).Width = 60
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                    userDate = userDate.AddDays(+7)
                    culture = System.Globalization.CultureInfo.CurrentCulture
                    intWeek = culture.Calendar.GetWeekOfYear(userDate, System.Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday)

                    _StrWeek = "Week- " & intWeek
                    _StrWeek1 = "Week " & intWeek
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns.Add(_StrWeek, _StrWeek1)
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).Group = agroup6
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).Width = 60
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right


                End If
                i = i + 1
            Next
            _WeekNo = 0
            '====================================================================================
            'Knitting
            i = 0
            vcWhere = "T15Sales_Order='" & Trim(cboSO.Text) & "' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "PROX"), New SqlParameter("@vcWhereClause1", vcWhere))
            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                Dim remain As Integer
                Dim noOfWeek As Integer
                Dim userDate As Date
                Dim _LastDate As Date
                Dim _TimeSpan As TimeSpan

                userDate = DateTime.Parse(M01.Tables(0).Rows(i)("T15Month") & "/1/" & M01.Tables(0).Rows(i)("T15Year"))
                ' MsgBox(WeekdayName(Weekday(userDate)))
                If WeekdayName(Weekday(userDate)) = "Sunday" Then
                    userDate = userDate.AddDays(-4)
                ElseIf WeekdayName(Weekday(userDate)) = "Monday" Then
                    userDate = userDate.AddDays(-5)
                ElseIf WeekdayName(Weekday(userDate)) = "Tuesday" Then
                    userDate = userDate.AddDays(-6)
                ElseIf WeekdayName(Weekday(userDate)) = "Thusday" Then
                    userDate = userDate.AddDays(-1)
                ElseIf WeekdayName(Weekday(userDate)) = "Friday" Then
                    userDate = userDate.AddDays(-2)
                ElseIf WeekdayName(Weekday(userDate)) = "Saturday" Then
                    userDate = userDate.AddDays(-3)

                End If


                _LastDate = DateTime.Parse(M01.Tables(0).Rows(i)("T15Month") & "/1/" & M01.Tables(0).Rows(i)("T15Year"))
                ' MsgBox(Date.DaysInMonth(_LastDate.Year, _LastDate.Month))
                _LastDate = DateTime.Parse(M01.Tables(0).Rows(i)("T15Month") & "/" & Date.DaysInMonth(_LastDate.Year, _LastDate.Month) & "/" & M01.Tables(0).Rows(i)("T15Year"))
                _TimeSpan = _LastDate.Subtract(userDate)
                _WeekNo = _TimeSpan.Days / 7
                i = i + 1
            Next

            agroup3 = UltraGrid2.DisplayLayout.Bands(0).Groups.Add("Group3")
            agroup3.Header.Caption = "Knitting"
            _WeekNo = _WeekNo * 60
            agroup3.Width = _WeekNo
            '=====================================================================================
            i = 0
            vcWhere = "T15Sales_Order='" & Trim(cboSO.Text) & "' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "PROX"), New SqlParameter("@vcWhereClause1", vcWhere))
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
                    ' userDate = userDate.AddDays(-1)
                ElseIf WeekdayName(Weekday(userDate)) = "Friday" Then
                    userDate = userDate.AddDays(-1)
                ElseIf WeekdayName(Weekday(userDate)) = "Saturday" Then
                    userDate = userDate.AddDays(-2)

                End If
                _LastDate = DateTime.Parse(M01.Tables(0).Rows(i)("T15Month") & "/1/" & M01.Tables(0).Rows(i)("T15Year"))
                ' MsgBox(Date.DaysInMonth(_LastDate.Year, _LastDate.Month))
                _LastDate = DateTime.Parse(M01.Tables(0).Rows(i)("T15Month") & "/" & Date.DaysInMonth(_LastDate.Year, _LastDate.Month) & "/" & M01.Tables(0).Rows(i)("T15Year"))
                _TimeSpan = _LastDate.Subtract(userDate)
                _WeekNo = _TimeSpan.Days / 7

                userDate = userDate.AddDays(+7)
                vcWhere = "T15Sales_Order='" & Trim(cboSO.Text) & "' AND t01bulk='1st Bulk' and T15Year=" & M01.Tables(0).Rows(i)("T15Year") & " and T15Month=" & M01.Tables(0).Rows(i)("T15Month") & ""
                T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "FSTB"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(T01) Then
                    If _1stBulk_Status = False Then
                        userDate = userDate.AddDays(-21)
                    Else
                        userDate = userDate.AddDays(-28)
                    End If
                Else
                    userDate = userDate.AddDays(-14)
                End If

                If _WeekNo = 5 Then
                    Dim culture As System.Globalization.CultureInfo
                    Dim intWeek As Integer
                    Dim _StrWeek As String

                    culture = System.Globalization.CultureInfo.CurrentCulture
                    intWeek = culture.Calendar.GetWeekOfYear(userDate, System.Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday)

                    _StrWeek = "Week " & intWeek
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns.Add(_StrWeek, _StrWeek)
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).Group = agroup3
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).Width = 60
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                    userDate = userDate.AddDays(+7)
                    culture = System.Globalization.CultureInfo.CurrentCulture
                    intWeek = culture.Calendar.GetWeekOfYear(userDate, System.Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday)

                    _StrWeek = "Week " & intWeek
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns.Add(_StrWeek, _StrWeek)
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).Group = agroup3
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).Width = 60
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                    userDate = userDate.AddDays(+7)
                    culture = System.Globalization.CultureInfo.CurrentCulture
                    intWeek = culture.Calendar.GetWeekOfYear(userDate, System.Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday)

                    _StrWeek = "Week " & intWeek
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns.Add(_StrWeek, _StrWeek)
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).Group = agroup3
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).Width = 60
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                    userDate = userDate.AddDays(+7)
                    culture = System.Globalization.CultureInfo.CurrentCulture
                    intWeek = culture.Calendar.GetWeekOfYear(userDate, System.Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday)

                    _StrWeek = "Week " & intWeek
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns.Add(_StrWeek, _StrWeek)
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).Group = agroup3
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).Width = 60
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                    userDate = userDate.AddDays(+7)
                    culture = System.Globalization.CultureInfo.CurrentCulture
                    intWeek = culture.Calendar.GetWeekOfYear(userDate, System.Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday)

                    _StrWeek = "Week " & intWeek
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns.Add(_StrWeek, _StrWeek)
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).Group = agroup3
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).Width = 60
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                ElseIf _WeekNo = 4 Then
                    Dim culture As System.Globalization.CultureInfo
                    Dim intWeek As Integer
                    Dim _StrWeek As String

                    culture = System.Globalization.CultureInfo.CurrentCulture
                    intWeek = culture.Calendar.GetWeekOfYear(userDate, System.Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday)

                    _StrWeek = "Week " & intWeek
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns.Add(_StrWeek, _StrWeek)
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).Group = agroup3
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).Width = 60
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                    userDate = userDate.AddDays(+7)
                    culture = System.Globalization.CultureInfo.CurrentCulture
                    intWeek = culture.Calendar.GetWeekOfYear(userDate, System.Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday)

                    _StrWeek = "Week " & intWeek
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns.Add(_StrWeek, _StrWeek)
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).Group = agroup3
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).Width = 60
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                    userDate = userDate.AddDays(+7)
                    culture = System.Globalization.CultureInfo.CurrentCulture
                    intWeek = culture.Calendar.GetWeekOfYear(userDate, System.Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday)

                    _StrWeek = "Week " & intWeek
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns.Add(_StrWeek, _StrWeek)
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).Group = agroup3
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).Width = 60
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                    userDate = userDate.AddDays(+7)
                    culture = System.Globalization.CultureInfo.CurrentCulture
                    intWeek = culture.Calendar.GetWeekOfYear(userDate, System.Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday)

                    _StrWeek = "Week " & intWeek
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns.Add(_StrWeek, _StrWeek)
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).Group = agroup3
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).Width = 60
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right


                End If
                i = i + 1
            Next
            '--------------------------------------------------------------------------
            'DYEING
            agroup4 = UltraGrid2.DisplayLayout.Bands(0).Groups.Add("Group4")
            agroup4.Header.Caption = "Dyeing"
            _WeekNo = _WeekNo * 60
            agroup4.Width = _WeekNo
            i = 0
            vcWhere = "T15Sales_Order='" & Trim(cboSO.Text) & "' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "PROX"), New SqlParameter("@vcWhereClause1", vcWhere))
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
                    ' userDate = userDate.AddDays(-1)
                ElseIf WeekdayName(Weekday(userDate)) = "Friday" Then
                    userDate = userDate.AddDays(-1)
                ElseIf WeekdayName(Weekday(userDate)) = "Saturday" Then
                    userDate = userDate.AddDays(-2)

                End If
                _LastDate = DateTime.Parse(M01.Tables(0).Rows(i)("T15Month") & "/1/" & M01.Tables(0).Rows(i)("T15Year"))
                ' MsgBox(Date.DaysInMonth(_LastDate.Year, _LastDate.Month))
                _LastDate = DateTime.Parse(M01.Tables(0).Rows(i)("T15Month") & "/" & Date.DaysInMonth(_LastDate.Year, _LastDate.Month) & "/" & M01.Tables(0).Rows(i)("T15Year"))
                _TimeSpan = _LastDate.Subtract(userDate)
                _WeekNo = _TimeSpan.Days / 7

                userDate = userDate.AddDays(+7)
                vcWhere = "T15Sales_Order='" & Trim(cboSO.Text) & "' AND t01bulk='1st Bulk' and T15Year=" & M01.Tables(0).Rows(i)("T15Year") & " and T15Month=" & M01.Tables(0).Rows(i)("T15Month") & ""
                T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "FSTB"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(T01) Then
                    userDate = userDate.AddDays(-21)
                Else
                    userDate = userDate.AddDays(-14)
                End If
                ' userDate = userDate.AddDays(+7)
                If _WeekNo = 5 Then
                    Dim culture As System.Globalization.CultureInfo
                    Dim intWeek As Integer
                    Dim _StrWeek As String

                    culture = System.Globalization.CultureInfo.CurrentCulture
                    intWeek = culture.Calendar.GetWeekOfYear(userDate, System.Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday)

                    _StrWeek = "Week. " & intWeek
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns.Add(_StrWeek, _StrWeek)
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).Group = agroup4
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).Width = 60
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                    userDate = userDate.AddDays(+7)
                    culture = System.Globalization.CultureInfo.CurrentCulture
                    intWeek = culture.Calendar.GetWeekOfYear(userDate, System.Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday)

                    _StrWeek = "Week. " & intWeek
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns.Add(_StrWeek, _StrWeek)
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).Group = agroup4
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).Width = 60
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                    userDate = userDate.AddDays(+7)
                    culture = System.Globalization.CultureInfo.CurrentCulture
                    intWeek = culture.Calendar.GetWeekOfYear(userDate, System.Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday)

                    _StrWeek = "Week. " & intWeek
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns.Add(_StrWeek, _StrWeek)
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).Group = agroup4
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).Width = 60
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                    userDate = userDate.AddDays(+7)
                    culture = System.Globalization.CultureInfo.CurrentCulture
                    intWeek = culture.Calendar.GetWeekOfYear(userDate, System.Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday)

                    _StrWeek = "Week. " & intWeek
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns.Add(_StrWeek, _StrWeek)
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).Group = agroup4
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).Width = 60
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                    userDate = userDate.AddDays(+7)
                    culture = System.Globalization.CultureInfo.CurrentCulture
                    intWeek = culture.Calendar.GetWeekOfYear(userDate, System.Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday)

                    _StrWeek = "Week. " & intWeek
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns.Add(_StrWeek, _StrWeek)
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).Group = agroup4
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).Width = 60
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                ElseIf _WeekNo = 4 Then
                    Dim culture As System.Globalization.CultureInfo
                    Dim intWeek As Integer
                    Dim _StrWeek As String

                    culture = System.Globalization.CultureInfo.CurrentCulture
                    intWeek = culture.Calendar.GetWeekOfYear(userDate, System.Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday)

                    _StrWeek = "Week. " & intWeek
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns.Add(_StrWeek, _StrWeek)
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).Group = agroup4
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).Width = 60
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                    userDate = userDate.AddDays(+7)
                    culture = System.Globalization.CultureInfo.CurrentCulture
                    intWeek = culture.Calendar.GetWeekOfYear(userDate, System.Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday)

                    _StrWeek = "Week. " & intWeek
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns.Add(_StrWeek, _StrWeek)
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).Group = agroup4
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).Width = 60
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                    userDate = userDate.AddDays(+7)
                    culture = System.Globalization.CultureInfo.CurrentCulture
                    intWeek = culture.Calendar.GetWeekOfYear(userDate, System.Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday)

                    _StrWeek = "Week. " & intWeek
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns.Add(_StrWeek, _StrWeek)
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).Group = agroup4
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).Width = 60
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                    userDate = userDate.AddDays(+7)
                    culture = System.Globalization.CultureInfo.CurrentCulture
                    intWeek = culture.Calendar.GetWeekOfYear(userDate, System.Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday)

                    _StrWeek = "Week. " & intWeek
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns.Add(_StrWeek, _StrWeek)
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).Group = agroup4
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).Width = 60
                    Me.UltraGrid2.DisplayLayout.Bands(0).Columns(_StrWeek).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right


                End If
                i = i + 1
            Next
            '-------------------------------------------------------------------------------->>>
            'DATA FILL YARN DYE
            Dim _Coloum_Count2 As Integer
            vcWhere = "select * from P01PARAMETER where P01CODE='PRN'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, vcWhere)
            If isValidDataset(M01) Then
                _Code = M01.Tables(0).Rows(0)("P01NO")
            End If

            _Code = _Code - 1
            _Row_Count = 0
            _Coloum_Count2 = 2
            _projection_Qty = 0
            vcWhere = "T15Sales_Order='" & Trim(cboSO.Text) & "' AND LEFT(M01Quality_No,1)='Y'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "PROV"), New SqlParameter("@vcWhereClause1", vcWhere))
            X = 0
            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                _Coloum_Count2 = 2
                _WeekNo = 0
                i = 0
                '  _projection_Qty = M01.Tables(0).Rows(i)("Qty")
                vcWhere = "T15Sales_Order='" & Trim(cboSO.Text) & "' AND LEFT(M01Quality_No,1)='Y' "
                T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "PROX"), New SqlParameter("@vcWhereClause1", vcWhere))
                For Each DTRow4 As DataRow In T01.Tables(0).Rows
                    Dim remain As Integer
                    Dim noOfWeek As Integer
                    Dim userDate As Date
                    Dim _LastDate As Date
                    Dim _TimeSpan As TimeSpan
                    Dim Z As Integer
                    Dim _week As Integer
                    Dim Value As Double
                    Dim _St As String

                    userDate = DateTime.Parse(T01.Tables(0).Rows(i)("T15Month") & "/1/" & T01.Tables(0).Rows(i)("T15Year"))
                    ' MsgBox(WeekdayName(Weekday(userDate)))
                    If WeekdayName(Weekday(userDate)) = "Sunday" Then
                        userDate = userDate.AddDays(-3)
                    ElseIf WeekdayName(Weekday(userDate)) = "Monday" Then
                        userDate = userDate.AddDays(-4)
                    ElseIf WeekdayName(Weekday(userDate)) = "Tuesday" Then
                        userDate = userDate.AddDays(-5)
                    ElseIf WeekdayName(Weekday(userDate)) = "Thusday" Then
                        ' userDate = userDate.AddDays(-1)
                    ElseIf WeekdayName(Weekday(userDate)) = "Friday" Then
                        userDate = userDate.AddDays(-1)
                    ElseIf WeekdayName(Weekday(userDate)) = "Saturday" Then
                        userDate = userDate.AddDays(-2)

                    End If
                    _LastDate = DateTime.Parse(T01.Tables(0).Rows(i)("T15Month") & "/1/" & T01.Tables(0).Rows(i)("T15Year"))
                    ' MsgBox(Date.DaysInMonth(_LastDate.Year, _LastDate.Month))
                    _LastDate = DateTime.Parse(T01.Tables(0).Rows(i)("T15Month") & "/" & Date.DaysInMonth(_LastDate.Year, _LastDate.Month) & "/" & T01.Tables(0).Rows(i)("T15Year"))
                    _TimeSpan = _LastDate.Subtract(userDate)
                    _WeekNo = _WeekNo + (_TimeSpan.Days / 7)
                    _week = (_TimeSpan.Days / 7)

                    vcWhere = "Code='" & Trim(M01.Tables(0).Rows(X)("T15Code")) & "' AND M43Shade='" & Trim(M01.Tables(0).Rows(X)("T15Shade")) & "' AND M43Count_No=" & _Code & " AND M43Product_Month=" & T01.Tables(0).Rows(i)("T15Month") & " AND M43Year=" & T01.Tables(0).Rows(i)("T15Year") & ""
                    dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "PRSX"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(dsUser) Then
                        Value = dsUser.Tables(0).Rows(0)("Qty") / _week
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

                        For Z = 1 To _week
                            If Trim(UltraGrid2.Rows(_Row_Count).Cells(_Coloum_Count2).Text) = "-" Then
                            Else
                                'Dim culture As System.Globalization.CultureInfo
                                'Dim intWeek As Integer
                                'Dim _StrWeek As String

                                'culture = System.Globalization.CultureInfo.CurrentCulture
                                'intWeek = culture.Calendar.GetWeekOfYear(userDate, System.Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday)

                                UltraGrid2.Rows(_Row_Count).Cells(_Coloum_Count2).Value = _St
                                userDate = userDate.AddDays(+7)
                                _Coloum_Count2 = _Coloum_Count2 + 1
                            End If
                        Next
                    Else
                        For Z = 1 To _week
                            If Trim(UltraGrid2.Rows(_Row_Count).Cells(_Coloum_Count2).Text) = "-" Then
                            Else
                                UltraGrid2.Rows(_Row_Count).Cells(_Coloum_Count2).Value = "-"
                                UltraGrid2.Rows(_Row_Count + 1).Cells(_Coloum_Count2).Value = "-"
                                _Coloum_Count2 = _Coloum_Count2 + 1
                            End If
                        Next
                    End If
                    i = i + 1
                Next

                _Row_Count = _Row_Count + 4
                X = X + 1
            Next


            '=================================================================================================
            'DATA FILL KNITTING
            '_Code = _Code - 1
            _Row_Count = 0
            _Coloum_Count = _Coloum_Count2
            _projection_Qty = 0



            vcWhere = "T15Sales_Order='" & Trim(cboSO.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "PROV"), New SqlParameter("@vcWhereClause1", vcWhere))
            X = 0
            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                _Coloum_Count = _Coloum_Count2
                _WeekNo = 0
                i = 0
                '  _projection_Qty = M01.Tables(0).Rows(i)("Qty")
                vcWhere = "T15Sales_Order='" & Trim(cboSO.Text) & "' "
                T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "PROX"), New SqlParameter("@vcWhereClause1", vcWhere))
                For Each DTRow4 As DataRow In T01.Tables(0).Rows
                    Dim remain As Integer
                    Dim noOfWeek As Integer
                    Dim userDate As Date
                    Dim _LastDate As Date
                    Dim _TimeSpan As TimeSpan
                    Dim Z As Integer
                    Dim _week As Integer
                    Dim Value As Double
                    Dim _St As String

                    userDate = DateTime.Parse(T01.Tables(0).Rows(i)("T15Month") & "/1/" & T01.Tables(0).Rows(i)("T15Year"))
                    ' MsgBox(WeekdayName(Weekday(userDate)))
                    If WeekdayName(Weekday(userDate)) = "Sunday" Then
                        userDate = userDate.AddDays(-3)
                    ElseIf WeekdayName(Weekday(userDate)) = "Monday" Then
                        userDate = userDate.AddDays(-4)
                    ElseIf WeekdayName(Weekday(userDate)) = "Tuesday" Then
                        userDate = userDate.AddDays(-5)
                    ElseIf WeekdayName(Weekday(userDate)) = "Thusday" Then
                        ' userDate = userDate.AddDays(-1)
                    ElseIf WeekdayName(Weekday(userDate)) = "Friday" Then
                        userDate = userDate.AddDays(-1)
                    ElseIf WeekdayName(Weekday(userDate)) = "Saturday" Then
                        userDate = userDate.AddDays(-2)

                    End If
                    _LastDate = DateTime.Parse(T01.Tables(0).Rows(i)("T15Month") & "/1/" & T01.Tables(0).Rows(i)("T15Year"))
                    ' MsgBox(Date.DaysInMonth(_LastDate.Year, _LastDate.Month))
                    _LastDate = DateTime.Parse(T01.Tables(0).Rows(i)("T15Month") & "/" & Date.DaysInMonth(_LastDate.Year, _LastDate.Month) & "/" & T01.Tables(0).Rows(i)("T15Year"))
                    _TimeSpan = _LastDate.Subtract(userDate)
                    _WeekNo = _WeekNo + (_TimeSpan.Days / 7)
                    _week = (_TimeSpan.Days / 7)

                    vcWhere = "Code='" & Trim(M01.Tables(0).Rows(X)("T15Code")) & "' AND M43Shade='" & Trim(M01.Tables(0).Rows(X)("T15Shade")) & "' AND M43Count_No=" & _Code & " AND M43Product_Month=" & T01.Tables(0).Rows(i)("T15Month") & " AND M43Year=" & T01.Tables(0).Rows(i)("T15Year") & ""
                    dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "PRSX"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(dsUser) Then
                        Value = dsUser.Tables(0).Rows(0)("Qty") / _week
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

                        For Z = 1 To _week
                            If Trim(UltraGrid2.Rows(_Row_Count).Cells(_Coloum_Count).Text) = "-" Then
                            Else
                                'Dim culture As System.Globalization.CultureInfo
                                'Dim intWeek As Integer
                                'Dim _StrWeek As String

                                'culture = System.Globalization.CultureInfo.CurrentCulture
                                'intWeek = culture.Calendar.GetWeekOfYear(userDate, System.Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday)

                                UltraGrid2.Rows(_Row_Count).Cells(_Coloum_Count).Value = _St
                                userDate = userDate.AddDays(+7)
                                _Coloum_Count = _Coloum_Count + 1
                            End If
                        Next
                    Else
                        For Z = 1 To _week
                            If Trim(UltraGrid2.Rows(_Row_Count).Cells(_Coloum_Count).Text) = "-" Then
                            Else
                                UltraGrid2.Rows(_Row_Count).Cells(_Coloum_Count).Value = "-"
                                UltraGrid2.Rows(_Row_Count + 1).Cells(_Coloum_Count).Value = "-"
                                _Coloum_Count = _Coloum_Count + 1
                            End If
                        Next
                    End If
                    i = i + 1
                Next

                _Row_Count = _Row_Count + 4
                X = X + 1
            Next

            '-------------------------------------------------------------------------------------------------------
            'Data Filling DYEING
            Dim _Coloum_Count1 As Integer
            _Row_Count = 0
            _Coloum_Count1 = _Coloum_Count
            vcWhere = "T15Sales_Order='" & Trim(cboSO.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "PROV"), New SqlParameter("@vcWhereClause1", vcWhere))
            X = 0
            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                _Coloum_Count1 = _Coloum_Count
                _WeekNo = 0
                i = 0
                vcWhere = "T15Sales_Order='" & Trim(cboSO.Text) & "' "
                T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "PROX"), New SqlParameter("@vcWhereClause1", vcWhere))
                For Each DTRow4 As DataRow In T01.Tables(0).Rows
                    Dim remain As Integer
                    Dim noOfWeek As Integer
                    Dim userDate As Date
                    Dim _LastDate As Date
                    Dim _TimeSpan As TimeSpan
                    Dim Z As Integer
                    Dim _week As Integer
                    Dim Value As Double
                    Dim _St As String

                    userDate = DateTime.Parse(T01.Tables(0).Rows(i)("T15Month") & "/1/" & T01.Tables(0).Rows(i)("T15Year"))
                    ' MsgBox(WeekdayName(Weekday(userDate)))
                    If WeekdayName(Weekday(userDate)) = "Sunday" Then
                        userDate = userDate.AddDays(-3)
                    ElseIf WeekdayName(Weekday(userDate)) = "Monday" Then
                        userDate = userDate.AddDays(-4)
                    ElseIf WeekdayName(Weekday(userDate)) = "Tuesday" Then
                        userDate = userDate.AddDays(-5)
                    ElseIf WeekdayName(Weekday(userDate)) = "Thusday" Then
                        ' userDate = userDate.AddDays(-1)
                    ElseIf WeekdayName(Weekday(userDate)) = "Friday" Then
                        userDate = userDate.AddDays(-1)
                    ElseIf WeekdayName(Weekday(userDate)) = "Saturday" Then
                        userDate = userDate.AddDays(-2)

                    End If
                    _LastDate = DateTime.Parse(T01.Tables(0).Rows(i)("T15Month") & "/1/" & T01.Tables(0).Rows(i)("T15Year"))
                    ' MsgBox(Date.DaysInMonth(_LastDate.Year, _LastDate.Month))
                    _LastDate = DateTime.Parse(T01.Tables(0).Rows(i)("T15Month") & "/" & Date.DaysInMonth(_LastDate.Year, _LastDate.Month) & "/" & T01.Tables(0).Rows(i)("T15Year"))
                    _TimeSpan = _LastDate.Subtract(userDate)
                    _WeekNo = _WeekNo + (_TimeSpan.Days / 7)
                    _week = (_TimeSpan.Days / 7)

                    vcWhere = "Code='" & Trim(M01.Tables(0).Rows(X)("T15Code")) & "' AND M43Shade='" & Trim(M01.Tables(0).Rows(X)("T15Shade")) & "' AND M43Count_No=" & _Code & " AND M43Product_Month=" & T01.Tables(0).Rows(i)("T15Month") & " AND M43Year=" & T01.Tables(0).Rows(i)("T15Year") & ""
                    dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "PRSX"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(dsUser) Then
                        Value = dsUser.Tables(0).Rows(0)("Qty") / _week
                        _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                        For Z = 1 To _week
                            If Trim(UltraGrid2.Rows(_Row_Count).Cells(_Coloum_Count1).Text) = "-" Then
                            Else
                                UltraGrid2.Rows(_Row_Count).Cells(_Coloum_Count1).Value = _St
                                _Coloum_Count1 = _Coloum_Count1 + 1
                            End If
                        Next
                    Else
                        For Z = 1 To _week
                            If Trim(UltraGrid2.Rows(_Row_Count).Cells(_Coloum_Count1).Text) = "-" Then
                            Else
                                UltraGrid2.Rows(_Row_Count).Cells(_Coloum_Count1).Value = "-"
                                UltraGrid2.Rows(_Row_Count + 1).Cells(_Coloum_Count).Value = "-"
                                _Coloum_Count1 = _Coloum_Count1 + 1
                            End If
                        Next
                    End If
                    i = i + 1
                Next

                _Row_Count = _Row_Count + 4
                X = X + 1
            Next
            '--------------------------------------------------------------------------------------------
            'FUNCTION DATA FILLING YARN DYE
            _Coloum_Count2 = 2
            _Row_Count = 1

            i = 0
            vcWhere = "T15Sales_Order='" & Trim(cboSO.Text) & "' AND LEFT(M01Quality_No,1)='Y'"
            T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "PROX"), New SqlParameter("@vcWhereClause1", vcWhere))
            For Each DTRow4 As DataRow In T01.Tables(0).Rows
                Dim remain As Integer
                Dim noOfWeek As Integer
                Dim userDate As Date
                Dim _LastDate As Date
                Dim _TimeSpan As TimeSpan
                Dim Z As Integer
                Dim _week As Integer
                Dim Value As Double
                Dim _St As String

                userDate = DateTime.Parse(T01.Tables(0).Rows(i)("T15Month") & "/1/" & T01.Tables(0).Rows(i)("T15Year"))
                ' MsgBox(WeekdayName(Weekday(userDate)))
                If WeekdayName(Weekday(userDate)) = "Sunday" Then
                    userDate = userDate.AddDays(-3)
                ElseIf WeekdayName(Weekday(userDate)) = "Monday" Then
                    userDate = userDate.AddDays(-4)
                ElseIf WeekdayName(Weekday(userDate)) = "Tuesday" Then
                    userDate = userDate.AddDays(-5)
                ElseIf WeekdayName(Weekday(userDate)) = "Thusday" Then
                    ' userDate = userDate.AddDays(-1)
                ElseIf WeekdayName(Weekday(userDate)) = "Friday" Then
                    userDate = userDate.AddDays(-1)
                ElseIf WeekdayName(Weekday(userDate)) = "Saturday" Then
                    userDate = userDate.AddDays(-2)

                End If


                _LastDate = DateTime.Parse(T01.Tables(0).Rows(i)("T15Month") & "/1/" & T01.Tables(0).Rows(i)("T15Year"))
                ' MsgBox(Date.DaysInMonth(_LastDate.Year, _LastDate.Month))
                _LastDate = DateTime.Parse(T01.Tables(0).Rows(i)("T15Month") & "/" & Date.DaysInMonth(_LastDate.Year, _LastDate.Month) & "/" & T01.Tables(0).Rows(i)("T15Year"))
                _TimeSpan = _LastDate.Subtract(userDate)
                _WeekNo = _WeekNo + (_TimeSpan.Days / 7)
                _week = (_TimeSpan.Days / 7)

                'CALCULATION TOTAL PROJECTION
                _projection_Qty = 0

                vcWhere = "T15Sales_Order='" & Trim(cboSO.Text) & "' AND T15Year=" & T01.Tables(0).Rows(i)("T15Year") & " AND T15Month=" & T01.Tables(0).Rows(i)("T15Month") & " "
                dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TPJQ"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(dsUser) Then
                    _projection_Qty = dsUser.Tables(0).Rows(0)("QTY")
                End If

                userDate = userDate.AddDays(+7)
                vcWhere = "T15Sales_Order='" & Trim(cboSO.Text) & "' AND t01bulk='1st Bulk' and T15Year=" & T01.Tables(0).Rows(i)("T15Year") & " and T15Month=" & T01.Tables(0).Rows(i)("T15Month") & ""
                T02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "FSTB"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(T02) Then
                    userDate = userDate.AddDays(-35)
                Else
                    userDate = userDate.AddDays(-28)
                End If

                '//--------------------------------------------------------------
                '  _Coloum_Count2 = 2
                For Z = 1 To _week
                    Dim culture As System.Globalization.CultureInfo
                    Dim intWeek As Integer
                    Dim _StrWeek As String

                    culture = System.Globalization.CultureInfo.CurrentCulture
                    intWeek = culture.Calendar.GetWeekOfYear(userDate, System.Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday)

                    X = 0
                    _Row_Count = 1
                    vcWhere = "T15Sales_Order='" & Trim(cboSO.Text) & "' AND LEFT(M01Quality_No,1)='Y'"
                    dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "PROV"), New SqlParameter("@vcWhereClause1", vcWhere))
                    For Each DTRow5 As DataRow In dsUser.Tables(0).Rows



                        vcWhere = "t15CODE='" & Trim(dsUser.Tables(0).Rows(X)("t15CODE")) & "' and T15SHADE='" & Trim(dsUser.Tables(0).Rows(X)("T15SHADE")) & "' and tmpWeek_No=" & intWeek & " and tmpYear=" & Year(userDate) & " AND LEFT(T15Quality,1)='Y'"
                        T02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "PKNG"), New SqlParameter("@vcWhereClause1", vcWhere))
                        If isValidDataset(T02) Then
                            If _projection_Qty >= T02.Tables(0).Rows(0)("Qty") Then
                                Value = T02.Tables(0).Rows(0)("Qty")
                                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                                UltraGrid2.Rows(_Row_Count).Cells(_Coloum_Count).Value = _St
                                _projection_Qty = _projection_Qty - T02.Tables(0).Rows(0)("Qty")

                                Value = CDbl(UltraGrid2.Rows(_Row_Count - 1).Cells(_Coloum_Count).Value)
                                Value = Value - T02.Tables(0).Rows(0)("Qty")
                                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                                UltraGrid2.Rows(_Row_Count + 1).Cells(_Coloum_Count).Value = _St
                                If Value > 0 Then
                                    UltraGrid2.Rows(_Row_Count + 1).Cells(_Coloum_Count).Appearance.BackColor = Color.LightGreen
                                End If
                                If X = 0 Then
                                    UltraGrid2.Rows(_Row_Count + 4).Cells(_Coloum_Count).Value = "-"
                                    UltraGrid2.Rows(_Row_Count + 5).Cells(_Coloum_Count).Value = UltraGrid2.Rows(_Row_Count + 3).Cells(_Coloum_Count).Value
                                    If UltraGrid2.Rows(_Row_Count + 3).Cells(_Coloum_Count).Value = "-" Then
                                    Else
                                        UltraGrid2.Rows(_Row_Count + 5).Cells(_Coloum_Count).Appearance.BackColor = Color.LightGreen
                                    End If
                                Else
                                    ' UltraGrid2.Rows(_Row_Count).Cells(_Coloum_Count).Value = "-"
                                End If
                                ' _Coloum_Count = _Coloum_Count + 1
                            Else
                                Value = _projection_Qty
                                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                                UltraGrid2.Rows(_Row_Count).Cells(_Coloum_Count).Value = _St
                                ' _Coloum_Count = _Coloum_Count + 1
                                i = i + 1
                                Exit For
                            End If
                        Else
                            UltraGrid2.Rows(_Row_Count).Cells(_Coloum_Count2).Value = "-"
                            UltraGrid2.Rows(_Row_Count + 1).Cells(_Coloum_Count2).Value = UltraGrid2.Rows(_Row_Count - 1).Cells(_Coloum_Count2).Value
                            If UltraGrid2.Rows(_Row_Count - 1).Cells(_Coloum_Count2).Value = "-" Then
                            Else
                                UltraGrid2.Rows(_Row_Count + 1).Cells(_Coloum_Count2).Appearance.BackColor = Color.LightGreen
                            End If
                            If X = 0 Then
                                UltraGrid2.Rows(_Row_Count + 4).Cells(_Coloum_Count2).Value = "-"
                            Else
                                ' UltraGrid2.Rows(_Row_Count - 4).Cells(_Coloum_Count).Value = "-"
                            End If
                            ' _Coloum_Count = _Coloum_Count + 1
                        End If


                        '_Coloum_Count = _Coloum_Count + 1
                        _Row_Count = _Row_Count + 4
                        X = X + 1
                    Next

                    _Coloum_Count2 = _Coloum_Count2 + 1
                    userDate = userDate.AddDays(+7)
                Next


                i = i + 1
            Next
            '============================================================================================


            _Coloum_Count = _Coloum_Count2
            'Function Data Filing Knitting
            _Row_Count = 1

            i = 0
            vcWhere = "T15Sales_Order='" & Trim(cboSO.Text) & "' "
            T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "PROX"), New SqlParameter("@vcWhereClause1", vcWhere))
            For Each DTRow4 As DataRow In T01.Tables(0).Rows
                Dim remain As Integer
                Dim noOfWeek As Integer
                Dim userDate As Date
                Dim _LastDate As Date
                Dim _TimeSpan As TimeSpan
                Dim Z As Integer
                Dim _week As Integer
                Dim Value As Double
                Dim _St As String

                userDate = DateTime.Parse(T01.Tables(0).Rows(i)("T15Month") & "/1/" & T01.Tables(0).Rows(i)("T15Year"))
                ' MsgBox(WeekdayName(Weekday(userDate)))
                If WeekdayName(Weekday(userDate)) = "Sunday" Then
                    userDate = userDate.AddDays(-3)
                ElseIf WeekdayName(Weekday(userDate)) = "Monday" Then
                    userDate = userDate.AddDays(-4)
                ElseIf WeekdayName(Weekday(userDate)) = "Tuesday" Then
                    userDate = userDate.AddDays(-5)
                ElseIf WeekdayName(Weekday(userDate)) = "Thusday" Then
                    ' userDate = userDate.AddDays(-1)
                ElseIf WeekdayName(Weekday(userDate)) = "Friday" Then
                    userDate = userDate.AddDays(-1)
                ElseIf WeekdayName(Weekday(userDate)) = "Saturday" Then
                    userDate = userDate.AddDays(-2)

                End If


                _LastDate = DateTime.Parse(T01.Tables(0).Rows(i)("T15Month") & "/1/" & T01.Tables(0).Rows(i)("T15Year"))
                ' MsgBox(Date.DaysInMonth(_LastDate.Year, _LastDate.Month))
                _LastDate = DateTime.Parse(T01.Tables(0).Rows(i)("T15Month") & "/" & Date.DaysInMonth(_LastDate.Year, _LastDate.Month) & "/" & T01.Tables(0).Rows(i)("T15Year"))
                _TimeSpan = _LastDate.Subtract(userDate)
                _WeekNo = _WeekNo + (_TimeSpan.Days / 7)
                _week = (_TimeSpan.Days / 7)

                'CALCULATION TOTAL PROJECTION
                _projection_Qty = 0

                vcWhere = "T15Sales_Order='" & Trim(cboSO.Text) & "' AND T15Year=" & T01.Tables(0).Rows(i)("T15Year") & " AND T15Month=" & T01.Tables(0).Rows(i)("T15Month") & ""
                dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TPJQ"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(dsUser) Then
                    _projection_Qty = dsUser.Tables(0).Rows(0)("QTY")
                End If

                userDate = userDate.AddDays(+7)
                vcWhere = "T15Sales_Order='" & Trim(cboSO.Text) & "' AND t01bulk='1st Bulk' and T15Year=" & T01.Tables(0).Rows(i)("T15Year") & " and T15Month=" & T01.Tables(0).Rows(i)("T15Month") & ""
                T02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "FSTB"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(T02) Then
                    userDate = userDate.AddDays(-21)
                Else
                    userDate = userDate.AddDays(-14)
                End If

                '//--------------------------------------------------------------
                ' _Coloum_Count = _Coloum_Count2
                For Z = 1 To _week
                    Dim culture As System.Globalization.CultureInfo
                    Dim intWeek As Integer
                    Dim _StrWeek As String

                    culture = System.Globalization.CultureInfo.CurrentCulture
                    intWeek = culture.Calendar.GetWeekOfYear(userDate, System.Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday)

                    X = 0
                    _Row_Count = 1
                    vcWhere = "T15Sales_Order='" & Trim(cboSO.Text) & "'"
                    dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "PROV"), New SqlParameter("@vcWhereClause1", vcWhere))
                    For Each DTRow5 As DataRow In dsUser.Tables(0).Rows



                        vcWhere = "t15CODE='" & Trim(dsUser.Tables(0).Rows(X)("t15CODE")) & "' and T15SHADE='" & Trim(dsUser.Tables(0).Rows(X)("T15SHADE")) & "' and tmpWeek_No=" & intWeek & " and tmpYear=" & Year(userDate) & ""
                        T02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "PKNG"), New SqlParameter("@vcWhereClause1", vcWhere))
                        If isValidDataset(T02) Then
                            If _projection_Qty >= T02.Tables(0).Rows(0)("Qty") Then
                                Value = T02.Tables(0).Rows(0)("Qty")
                                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                                UltraGrid2.Rows(_Row_Count).Cells(_Coloum_Count).Value = _St
                                _projection_Qty = _projection_Qty - T02.Tables(0).Rows(0)("Qty")

                                Value = CDbl(UltraGrid2.Rows(_Row_Count - 1).Cells(_Coloum_Count).Value)
                                Value = Value - T02.Tables(0).Rows(0)("Qty")
                                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                                UltraGrid2.Rows(_Row_Count + 1).Cells(_Coloum_Count).Value = _St
                                If Value > 0 Then
                                    UltraGrid2.Rows(_Row_Count + 1).Cells(_Coloum_Count).Appearance.BackColor = Color.LightGreen
                                End If
                                If X = 0 Then
                                    UltraGrid2.Rows(_Row_Count + 4).Cells(_Coloum_Count).Value = "-"
                                    UltraGrid2.Rows(_Row_Count + 5).Cells(_Coloum_Count).Value = UltraGrid2.Rows(_Row_Count + 3).Cells(_Coloum_Count).Value
                                    If UltraGrid2.Rows(_Row_Count + 3).Cells(_Coloum_Count).Value = "-" Then
                                    Else
                                        UltraGrid2.Rows(_Row_Count + 5).Cells(_Coloum_Count).Appearance.BackColor = Color.LightGreen
                                    End If
                                Else
                                    ' UltraGrid2.Rows(_Row_Count).Cells(_Coloum_Count).Value = "-"
                                End If
                                ' _Coloum_Count = _Coloum_Count + 1
                            Else
                                Value = _projection_Qty
                                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                                UltraGrid2.Rows(_Row_Count).Cells(_Coloum_Count).Value = _St
                                ' _Coloum_Count = _Coloum_Count + 1
                                i = i + 1
                                Exit For
                            End If
                        Else
                            UltraGrid2.Rows(_Row_Count).Cells(_Coloum_Count).Value = "-"
                            UltraGrid2.Rows(_Row_Count + 1).Cells(_Coloum_Count).Value = UltraGrid2.Rows(_Row_Count - 1).Cells(_Coloum_Count).Value
                            If UltraGrid2.Rows(_Row_Count - 1).Cells(_Coloum_Count).Value = "-" Then
                            Else
                                UltraGrid2.Rows(_Row_Count + 1).Cells(_Coloum_Count).Appearance.BackColor = Color.LightGreen
                            End If
                            If X = 0 Then
                                UltraGrid2.Rows(_Row_Count + 3).Cells(_Coloum_Count).Value = "-"
                            Else
                                ' UltraGrid2.Rows(_Row_Count - 4).Cells(_Coloum_Count).Value = "-"
                            End If
                            ' _Coloum_Count = _Coloum_Count + 1
                        End If


                        '_Coloum_Count = _Coloum_Count + 1
                        _Row_Count = _Row_Count + 4
                        X = X + 1
                    Next

                    _Coloum_Count = _Coloum_Count + 1
                    userDate = userDate.AddDays(+7)
                Next


                i = i + 1
            Next


            con.Close()
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.Close()
            End If
        End Try
    End Function


    Function Search_DY_KNTWeek(ByVal strCode As String)
        Dim M02 As DataSet
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Try
            Sql = "select * from T21Pack_Allocation where T21Sales_Order='" & strSales_Order & "'"
            M02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M02) Then
                With frmKnitting_Plan_WithTab
                    .lblP_YD.Text = M02.Tables(0).Rows(0)("T21Yan_Week")
                    .lblP_Knitting.Text = M02.Tables(0).Rows(0)("T21Knt_Week")
                    .lblP_Dye.Text = M02.Tables(0).Rows(0)("T21Dye_Week")

                    .lblY_Yd.Text = M02.Tables(0).Rows(0)("T21Yan_Week")
                    .lblY_Knt.Text = M02.Tables(0).Rows(0)("T21Knt_Week")
                    .lblY_Dye.Text = M02.Tables(0).Rows(0)("T21Dye_Week")

                    .lblB_Yd.Text = M02.Tables(0).Rows(0)("T21Yan_Week")
                    .lblB_Knt.Text = M02.Tables(0).Rows(0)("T21Knt_Week")
                    .lblB_Dye.Text = M02.Tables(0).Rows(0)("T21Dye_Week")

                    .lblK_Yd.Text = M02.Tables(0).Rows(0)("T21Yan_Week")
                    .lblK_Knt.Text = M02.Tables(0).Rows(0)("T21Knt_Week")
                    .lblK_Dye.Text = M02.Tables(0).Rows(0)("T21Dye_Week")

                    .lblD_Yd.Text = M02.Tables(0).Rows(0)("T21Yan_Week")
                    .lblD_Knt.Text = M02.Tables(0).Rows(0)("T21Knt_Week")
                    .lblD_Dye.Text = M02.Tables(0).Rows(0)("T21Dye_Week")
                End With
            End If
            con.Close()

        Catch returnMessage As ExecutionEngineException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
              
            End If
        End Try
    End Function


    Private Sub txtPk_Knt_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPk_Knt.KeyUp
        If e.KeyCode = 13 Then
            txtPk_Dye.Focus()
        ElseIf e.KeyCode = Keys.Tab Then
            txtPk_Dye.Focus()
        End If
    End Sub

    Private Sub txtPk_Dye_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPk_Dye.KeyUp
        If e.KeyCode = 13 Then
            cmdEx.Focus()
        ElseIf e.KeyCode = Keys.Tab Then
            cmdEx.Focus()
        End If
    End Sub

 
   
    Private Sub cmdEx_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEx.Click
        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean
        Dim vcWhere As String

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Dim M01 As DataSet
        Dim ncQryType As String

        Try
            If IsNumeric(txtPk_Dye.Text) Then

            Else
                MsgBox("Please enter the correct Dye Week", MsgBoxStyle.Information, "Information .......")
                txtPk_Dye.Focus()
                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
                connection.Close()
                Exit Sub
            End If

            If IsNumeric(txtYarn_Week.Text) Then

            Else
                MsgBox("Please enter the correct Yarn Dye Week", MsgBoxStyle.Information, "Information .......")
                txtYarn_Week.Focus()
                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
                connection.Close()
                Exit Sub
            End If

            If IsNumeric(txtPk_Knt.Text) Then

            Else
                MsgBox("Please enter the correct Knitting Week", MsgBoxStyle.Information, "Information .......")
                txtPk_Knt.Focus()
                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
                connection.Close()
                Exit Sub
            End If

            If Len(Trim(txtPk_Dye.Text)) > 2 Then
                MsgBox("Please enter the correct Dye Week", MsgBoxStyle.Information, "Information .......")
                txtPk_Dye.Focus()
                Exit Sub
            End If


            If Len(Trim(txtPk_Knt.Text)) > 2 Then
                MsgBox("Please enter the correct Knitting Week", MsgBoxStyle.Information, "Information .......")
                txtPk_Knt.Focus()
                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
                connection.Close()
                Exit Sub
            End If

            If Len(Trim(txtYarn_Week.Text)) > 2 Then
                MsgBox("Please enter the correct Yarn Dye Week", MsgBoxStyle.Information, "Information .......")
                txtYarn_Week.Focus()
                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
                connection.Close()
                Exit Sub
            End If

            ncQryType = "IPD"
            vcWhere = "T21Sales_Order='" & cboSO.Text & "'"
            M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetT21Pack_Allocation", New SqlParameter("@cQryType", "MXD"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                nvcFieldList1 = "UPDATE T21Pack_Allocation SET T21Knt_Week=" & txtPk_Knt.Text & ",T21Dye_Week=" & txtPk_Dye.Text & ",T21Yan_Week=" & txtYarn_Week.Text & " WHERE T21Sales_Order='" & cboSO.Text & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            Else
                nvcFieldList1 = "(T21Sales_Order," & "T21Knt_Week," & "T21Dye_Week," & "T21Date," & "T21User," & "T21Yan_Week) " & "values('" & cboSO.Text & "'," & Trim(txtPk_Knt.Text) & ",'" & txtPk_Dye.Text & "','" & Now & "','" & strDisname & "','" & txtYarn_Week.Text & "')"
                up_GetSetT21Pack_Allocation(ncQryType, nvcFieldList1, vcWhere, connection, transaction)
            End If

            transaction.Commit()
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
            connection.Close()

        Catch returnMessage As ExecutionEngineException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
                connection.Close()
            End If
        End Try
    End Sub

    Private Sub txtYarn_Week_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtYarn_Week.KeyUp
        If e.KeyCode = 13 Then
            txtYarn_Week.Focus()
        ElseIf e.KeyCode = Keys.Tab Then
            txtYarn_Week.Focus()
        End If
    End Sub

 
    Private Sub UltraGrid2_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles UltraGrid2.InitializeLayout

    End Sub
End Class