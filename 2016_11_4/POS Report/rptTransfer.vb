Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports CrystalDecisions.CrystalReports.Engine
Public Class rptTransfer
    Dim c_dataCustomer1 As DataTable
    Dim c_dataCustomer2 As DataTable
    Dim _PrintStatus As String
    Dim _From As Date
    Dim _To As Date
    Dim _Itemcode As String
    Dim _Comcode As String
    Dim _Supplier As String
    Dim _Category As String

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Function Load_Gride()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTable_Transfer
        UltraGrid1.DataSource = c_dataCustomer1
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 70
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 190
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(3).Width = 90
            .DisplayLayout.Bands(0).Columns(3).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(4).Width = 130
            .DisplayLayout.Bands(0).Columns(4).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(5).Width = 70
            .DisplayLayout.Bands(0).Columns(5).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(6).Width = 70
            .DisplayLayout.Bands(0).Columns(6).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(7).Width = 70
            .DisplayLayout.Bands(0).Columns(7).AutoEdit = False
            '.DisplayLayout.Bands(0).Columns(8).Width = 70
            '.DisplayLayout.Bands(0).Columns(8).AutoEdit = False
            '.DisplayLayout.Bands(0).Columns(9).Width = 90
            '.DisplayLayout.Bands(0).Columns(9).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(9).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

        End With
    End Function

    Function Load_Item()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M03Item_Code as [##] from M03Item_Master where m03Status='A' and M03Location='" & _Comcode & "' order by M03Item_Code "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboItem
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 220
                ' .Rows.Band.Columns(1).Width = 180


            End With
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            'With txtNSL
            '    .DataSource = M01
            '    .Rows.Band.Columns(0).Width = 225
            'End With

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub rptTransfer_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        _Comcode = ConfigurationManager.AppSettings("LOCCode")
        Call Load_Gride()
        txtDate3.Text = Today
        txtDate4.Text = Today
        txtA1.Text = Today
        txtA2.Text = Today
        Call Load_Location()
        txtB1.Text = Today
        txtB2.Text = Today
        Call Load_Category()
        txtC1.Text = Today
        txtC2.Text = Today
        Call Load_Item()

    End Sub

    Function Load_Category()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M02Cat_Name as [##] from M02Category where M02Com_Code='" & _Comcode & "' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboCategory
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 140
                ' .Rows.Band.Columns(1).Width = 180


            End With

            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try

    End Function

    Function Load_Location()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M04Loc_Name as [##] from M04Location  "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboSupplier
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 210
                ' .Rows.Band.Columns(1).Width = 180


            End With

            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try

    End Function

    Private Sub ResetToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ResetToolStripMenuItem.Click
        Call Load_Gride()
        Panel1.Visible = False
        Panel2.Visible = False
        Panel3.Visible = False
        Panel4.Visible = False
    End Sub

    Private Sub UsingDateToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UsingDateToolStripMenuItem.Click
        Panel1.Visible = False
        Panel2.Visible = False
        Panel3.Visible = True
        Panel4.Visible = False
        _PrintStatus = "A1"
    End Sub

    Function Load_Gride1()
        Dim Sql As String

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
        Dim _Qty As Integer
        Dim Value As Double
        Dim _Rowcount As Integer
        Dim _St As String
        Try
            Sql = "select *  from View_Transfer1 inner join View_T02Transaction on T02Ref_no=T01Ref_no where t01date between '" & txtDate3.Text & "' and '" & txtDate4.Text & "' and T01FromLoc_code='" & _Comcode & "' order by T01Ref_No"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Qty = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                newRow("Transfer No") = M01.Tables(0).Rows(i)("T01grn_No")
                newRow("Date") = Month(M01.Tables(0).Rows(i)("T01Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T01Date")) & "/" & Year(M01.Tables(0).Rows(i)("T01Date"))
                'newRow("Com Invoice") = M01.Tables(0).Rows(i)("T01Invoice_No")
                newRow("To Location") = M01.Tables(0).Rows(i)("M04Loc_Name")
                newRow("Item Code") = M01.Tables(0).Rows(i)("T02Item_Code")
                newRow("Item Name") = M01.Tables(0).Rows(i)("M03Item_Name")
                Value = M01.Tables(0).Rows(i)("T02Qty")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Qty") = _St

                'Value = M01.Tables(0).Rows(i)("T02Free_Issue")
                '_St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                '_St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                'newRow("Free Issue") = _St
                Value = M01.Tables(0).Rows(i)("T02Cost")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Rate") = _St

                Value = M01.Tables(0).Rows(i)("T02total")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Total") = _St

                c_dataCustomer1.Rows.Add(newRow)

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

    Function Load_Gride_Item()
        Dim Sql As String

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
        Dim _Qty As Integer
        Dim Value As Double
        Dim _Rowcount As Integer
        Dim _St As String
        Try
            Sql = "select *  from View_Transfer1 inner join View_T02Transaction on T02Ref_no=T01Ref_no where t01date between '" & txtC1.Text & "' and '" & txtC2.Text & "' and T01FromLoc_code='" & _Comcode & "' and M03Item_Code='" & Trim(cboItem.Text) & "' order by T01Ref_No"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Qty = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                newRow("Transfer No") = M01.Tables(0).Rows(i)("T01grn_No")
                newRow("Date") = Month(M01.Tables(0).Rows(i)("T01Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T01Date")) & "/" & Year(M01.Tables(0).Rows(i)("T01Date"))
                'newRow("Com Invoice") = M01.Tables(0).Rows(i)("T01Invoice_No")
                newRow("To Location") = M01.Tables(0).Rows(i)("M04Loc_Name")
                newRow("Item Code") = M01.Tables(0).Rows(i)("T02Item_Code")
                newRow("Item Name") = M01.Tables(0).Rows(i)("M03Item_Name")
                Value = M01.Tables(0).Rows(i)("T02Qty")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Qty") = _St

                'Value = M01.Tables(0).Rows(i)("T02Free_Issue")
                '_St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                '_St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                'newRow("Free Issue") = _St
                Value = M01.Tables(0).Rows(i)("T02Cost")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Rate") = _St

                Value = M01.Tables(0).Rows(i)("T02total")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Total") = _St

                c_dataCustomer1.Rows.Add(newRow)

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


    Function Load_Gride_Location()
        Dim Sql As String

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
        Dim _Qty As Integer
        Dim Value As Double
        Dim _Rowcount As Integer
        Dim _St As String
        Try
            Sql = "select *  from View_Transfer1 inner join View_T02Transaction on T02Ref_no=T01Ref_no where t01date between '" & txtA1.Text & "' and '" & txtA2.Text & "' and T01FromLoc_code='" & _Comcode & "' and M04Loc_Name='" & Trim(cboSupplier.Text) & "' order by T01Ref_No"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Qty = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                newRow("Transfer No") = M01.Tables(0).Rows(i)("T01grn_No")
                newRow("Date") = Month(M01.Tables(0).Rows(i)("T01Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T01Date")) & "/" & Year(M01.Tables(0).Rows(i)("T01Date"))
                'newRow("Com Invoice") = M01.Tables(0).Rows(i)("T01Invoice_No")
                newRow("To Location") = M01.Tables(0).Rows(i)("M04Loc_Name")
                newRow("Item Code") = M01.Tables(0).Rows(i)("T02Item_Code")
                newRow("Item Name") = M01.Tables(0).Rows(i)("M03Item_Name")
                Value = M01.Tables(0).Rows(i)("T02Qty")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Qty") = _St

                'Value = M01.Tables(0).Rows(i)("T02Free_Issue")
                '_St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                '_St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                'newRow("Free Issue") = _St
                Value = M01.Tables(0).Rows(i)("T02Cost")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Rate") = _St

                Value = M01.Tables(0).Rows(i)("T02total")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Total") = _St

                c_dataCustomer1.Rows.Add(newRow)

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

    Function Load_Gride_Category()
        Dim Sql As String

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
        Dim M02 As DataSet
        Dim _Qty As Integer
        Dim Value As Double
        Dim _Rowcount As Integer
        Dim _St As String
        Try
            Sql = "select * from M02Category where M02Cat_Name='" & Trim(cboCategory.Text) & "' and M02Com_Code='" & _Comcode & "'"
            M02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M02) Then
                Sql = "select *  from View_Transfer1 inner join View_T02Transaction on T02Ref_no=T01Ref_no where t01date between '" & txtB1.Text & "' and '" & txtB2.Text & "' and T01FromLoc_code='" & _Comcode & "' and M03Cat_Code='" & M02.Tables(0).Rows(0)("M02Cat_Code") & "' order by T01Ref_No"
                M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                i = 0
                _Qty = 0
                For Each DTRow1 As DataRow In M01.Tables(0).Rows
                    Dim newRow As DataRow = c_dataCustomer1.NewRow

                    newRow("Transfer No") = M01.Tables(0).Rows(i)("T01grn_No")
                    newRow("Date") = Month(M01.Tables(0).Rows(i)("T01Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T01Date")) & "/" & Year(M01.Tables(0).Rows(i)("T01Date"))
                    'newRow("Com Invoice") = M01.Tables(0).Rows(i)("T01Invoice_No")
                    newRow("To Location") = M01.Tables(0).Rows(i)("M04Loc_Name")
                    newRow("Item Code") = M01.Tables(0).Rows(i)("T02Item_Code")
                    newRow("Item Name") = M01.Tables(0).Rows(i)("M03Item_Name")
                    Value = M01.Tables(0).Rows(i)("T02Qty")
                    _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                    newRow("Qty") = _St

                    'Value = M01.Tables(0).Rows(i)("T02Free_Issue")
                    '_St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    '_St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                    'newRow("Free Issue") = _St
                    Value = M01.Tables(0).Rows(i)("T02Cost")
                    _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    newRow("Rate") = _St

                    Value = M01.Tables(0).Rows(i)("T02total")
                    _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    newRow("Total") = _St

                    c_dataCustomer1.Rows.Add(newRow)

                    i = i + 1
                Next
            End If
            con.close()
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Private Sub UltraButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton2.Click
        Call Load_Gride()
        Call Load_Gride1()
        _From = txtDate3.Text
        _To = txtDate4.Text
        Panel1.Visible = False
        Panel2.Visible = False
        Panel3.Visible = False
        Panel4.Visible = False
    End Sub

    Private Sub PrintToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrintToolStripMenuItem.Click
        Dim B As New ReportDocument
        Dim A As String
        Dim StrFromDate As String
        Dim StrToDate As String
        Try
            StrFromDate = "(" & Year(_From) & ", " & VB6.Format(Month(_From), "0#") & ", " & VB6.Format(CDate(_From).Day, "0#") & ", 00, 00, 00)"
            StrToDate = "(" & Year(_To) & ", " & VB6.Format(Month(_To), "0#") & ", " & VB6.Format(CDate(_To).Day, "0#") & ", 00, 00, 00)"

            If _PrintStatus = "A1" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\Transfer.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                B.SetParameterValue("To", _To)
                B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_Transfer1.T01Date}  in DateTime " & StrFromDate & " to DateTime " & StrToDate & " and  {View_Transfer1.T01FromLoc_Code}='" & _Comcode & "'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _PrintStatus = "A2" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\Transfer.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                B.SetParameterValue("To", _To)
                B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_Transfer1.T01Date}  in DateTime " & StrFromDate & " to DateTime " & StrToDate & " and  {View_Transfer1.T01FromLoc_Code}='" & _Comcode & "' and {View_Transfer1.M04Loc_Name}='" & _Supplier & "'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _PrintStatus = "A3" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\Transfer.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                B.SetParameterValue("To", _To)
                B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_Transfer1.T01Date}  in DateTime " & StrFromDate & " to DateTime " & StrToDate & " and  {View_Transfer1.T01FromLoc_Code}='" & _Comcode & "' and {M02Category.M02Cat_Name}='" & _Category & "'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _PrintStatus = "A4" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\Transfer.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                B.SetParameterValue("To", _To)
                B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_Transfer1.T01Date}  in DateTime " & StrFromDate & " to DateTime " & StrToDate & " and  {View_Transfer1.T01FromLoc_Code}='" & _Comcode & "' and {View_T02Transaction.M03Item_Code}='" & _Itemcode & "'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            End If

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                ' con.close()
            End If
        End Try
    End Sub

    Private Sub UsingSupplierToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UsingSupplierToolStripMenuItem.Click
        Panel1.Visible = True
        Panel2.Visible = False
        Panel3.Visible = False
        Panel4.Visible = False
        _PrintStatus = "A2"
    End Sub

    Private Sub UltraButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton1.Click
        Call Load_Gride()
        _From = txtA1.Text
        _To = txtA2.Text
        _Supplier = Trim(cboSupplier.Text)
        Call Load_Gride_Location()
        Panel1.Visible = False
        Panel2.Visible = False
        Panel3.Visible = False
        Panel4.Visible = False
    End Sub

    Private Sub UsingCategoryToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UsingCategoryToolStripMenuItem.Click
        Panel1.Visible = False
        Panel2.Visible = True
        Panel3.Visible = False
        Panel4.Visible = False
        _PrintStatus = "A3"
    End Sub

    Private Sub UltraButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton3.Click
        Call Load_Gride()
        _From = txtB1.Text
        _To = txtB2.Text
        _Category = Trim(cboCategory.Text)
        Call Load_Gride_Category()
    End Sub

    Private Sub UsingItemNameToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UsingItemNameToolStripMenuItem.Click
        Panel1.Visible = False
        Panel2.Visible = False
        Panel3.Visible = False
        Panel4.Visible = True
        _PrintStatus = "A4"
    End Sub

    Private Sub UltraButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton4.Click
        Call Load_Gride()
        _From = txtC1.Text
        _To = txtC2.Text
        _Itemcode = Trim(cboItem.Text)
        Call Load_Gride_Item()
        Panel1.Visible = False
        Panel2.Visible = False
        Panel3.Visible = False
        Panel4.Visible = False
    End Sub

    Private Sub Panel1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel1.Paint

    End Sub
End Class