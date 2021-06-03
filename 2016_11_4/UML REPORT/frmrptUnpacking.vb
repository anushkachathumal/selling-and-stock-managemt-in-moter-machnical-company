Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports CrystalDecisions.CrystalReports.Engine
Public Class frmrptUnpacking
    Dim c_dataCustomer1 As DataTable
    Dim c_dataCustomer2 As DataTable
    Dim _PrintStatus As String
    Dim _Itemcode As String
    Dim _From As Date
    Dim _to As Date
    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Function Load_Gride_Production()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTable_ProductionItemUnpacking
        UltraGrid1.DataSource = c_dataCustomer1
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 90
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False

            .DisplayLayout.Bands(0).Columns(3).Width = 210
            .DisplayLayout.Bands(0).Columns(3).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(4).Width = 70
            .DisplayLayout.Bands(0).Columns(4).AutoEdit = False

            '.DisplayLayout.Bands(0).Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            '.DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            '.DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            '.DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

            '.DisplayLayout.Bands(0).Columns(1).CellAppearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            '.DisplayLayout.Bands(0).Columns(0).CellAppearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            ' .DisplayLayout.Bands(0).Columns(1).
            ' .DisplayLayout.Bands(0).Header.Height = 60

            '  .DisplayLayout.Bands(0).Columns(6).Width = 90
            ' .DisplayLayout.Bands(0).Columns(7).Width = 90

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function

    Function Load_Gride_Production_Detailes()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTable_ProductionItemUnpackingD
        UltraGrid1.DataSource = c_dataCustomer1
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 210
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False

            .DisplayLayout.Bands(0).Columns(3).Width = 150
            .DisplayLayout.Bands(0).Columns(3).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(4).Width = 70
            .DisplayLayout.Bands(0).Columns(4).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(5).Width = 70
            .DisplayLayout.Bands(0).Columns(5).AutoEdit = False

            '.DisplayLayout.Bands(0).Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            '.DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            '.DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

            '.DisplayLayout.Bands(0).Columns(1).CellAppearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            '.DisplayLayout.Bands(0).Columns(0).CellAppearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            ' .DisplayLayout.Bands(0).Columns(1).
            ' .DisplayLayout.Bands(0).Header.Height = 60

            '  .DisplayLayout.Bands(0).Columns(6).Width = 90
            ' .DisplayLayout.Bands(0).Columns(7).Width = 90

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function

    Private Sub frmrptUnpacking_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_Gride_Production()
        txtDate3.Text = Today
        txtDate4.Text = Today
        Call Load_Combo()
        txtDate.Text = Today
        txtDate1.Text = Today
        txtDate5.Text = Today
        txtDate6.Text = Today
    End Sub


    Function Search_Itemcode() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim M02 As DataSet
        Dim _qty As Integer
        Dim _stockIn As Integer
        Try
            Sql = "select * from View_Production_Items where M14Status='A' and M14Item_Name='" & cboItem.Text & "' and category='PS' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then

                _Itemcode = Trim(M01.Tables(0).Rows(0)("M14Item_code"))
                Search_Itemcode = True


            End If
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            con.close()
            'With txtNSL
            '    .DataSource = M01
            '    .Rows.Band.Columns(0).Width = 225
            'End With

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Private Sub ResetToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ResetToolStripMenuItem.Click
        Call Load_Gride_Production()
        Panel1.Visible = False
        Panel2.Visible = False
        Panel3.Visible = False
    End Sub

    Private Sub ProductionSetToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ProductionSetToolStripMenuItem.Click

    End Sub

    Private Sub DateByDateToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateByDateToolStripMenuItem.Click
        _PrintStatus = "A"
        Panel1.Visible = True
    End Sub

    Function Load_Data_Grid1()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim i As Integer

        Try
            _From = txtDate3.Text
            _to = txtDate4.Text

            Sql = "select * from T05Unpacking_Header inner join View_Production_Items on t05Pro_code=m14item_code where t05status='A' and t05date between '" & txtDate3.Text & "' and '" & txtDate4.Text & "' order by T05Ref_No"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                newRow("Unpacking No") = M01.Tables(0).Rows(i)("T05Ref_No")
                newRow("Date") = Month(M01.Tables(0).Rows(i)("T05Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T05Date")) & "/" & Year(M01.Tables(0).Rows(i)("T05Date"))
                newRow("Item Code") = M01.Tables(0).Rows(i)("m14Item_Code")
                newRow("Item Name") = M01.Tables(0).Rows(i)("M14Item_Name")
                newRow("Qty") = M01.Tables(0).Rows(i)("T05Qty")
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

    Function Load_Data_Grid2()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim i As Integer

        Try
            _From = txtDate.Text
            _to = txtDate1.Text
            Sql = "select * from T05Unpacking_Header inner join View_Production_Items on t05Pro_code=m14item_code where t05status='A' and t05date between '" & txtDate.Text & "' and '" & txtDate1.Text & "' and m14Item_Code='" & _Itemcode & "' order by T05Ref_No"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                newRow("Unpacking No") = M01.Tables(0).Rows(i)("T05Ref_No")
                newRow("Date") = Month(M01.Tables(0).Rows(i)("T05Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T05Date")) & "/" & Year(M01.Tables(0).Rows(i)("T05Date"))
                newRow("Item Code") = M01.Tables(0).Rows(i)("m14Item_Code")
                newRow("Item Name") = M01.Tables(0).Rows(i)("M14Item_Name")
                newRow("Qty") = M01.Tables(0).Rows(i)("T05Qty")
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

    Function Load_Data_Grid3()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim i As Integer

        Try
            _From = txtDate5.Text
            _to = txtDate6.Text

            Sql = "select T05Ref_No,T05Date,v.m14item_name,p.m14item_name as Item,T06Qty,T06Recycle from T05Unpacking_Header inner join View_Production_Items V on t05Pro_code=m14item_code inner join T06Unpacking_Fluter on T05Ref_No=T06Ref_No inner join M14Product_Item P on P.M14Item_Code=T06Item_Code where t05status='A' and t05date between '" & txtDate5.Text & "' and '" & txtDate6.Text & "'  order by T05Ref_No"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                newRow("Unpacking No") = M01.Tables(0).Rows(i)("T05Ref_No")
                newRow("Date") = Month(M01.Tables(0).Rows(i)("T05Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T05Date")) & "/" & Year(M01.Tables(0).Rows(i)("T05Date"))
                newRow("Set Name") = M01.Tables(0).Rows(i)("M14Item_Name")
                newRow("Item Name") = M01.Tables(0).Rows(i)("Item")
                newRow("Qty") = M01.Tables(0).Rows(i)("T06Qty")
                newRow("Recycle Qty") = M01.Tables(0).Rows(i)("T06Recycle")
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

    Private Sub UltraButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton1.Click
        Call Load_Gride_Production()
        _PrintStatus = "A"
        Call Load_Data_Grid1()
        Panel1.Visible = False
    End Sub

 
    Private Sub ByProductionSetToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ByProductionSetToolStripMenuItem.Click
        Call Load_Gride_Production()
        _PrintStatus = "B"
        ' Call Load_Data_Grid1()
        Panel2.Visible = True
    End Sub

    Function Load_Combo()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M14Item_Name as [##] from View_Production_Items where M14Status='A' and Category='PS' order by M14Item_Code "
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

    Private Sub UltraButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton2.Click
        If Search_Itemcode() = True Then
        Else
            MsgBox("Please select Item Name", MsgBoxStyle.Information, "Information ......")
            cboItem.ToggleDropdown()
            Exit Sub
        End If

        Call Load_Gride_Production()
        Call Load_Data_Grid2()
        Panel2.Visible = False
    End Sub

    Private Sub ProductionItemToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ProductionItemToolStripMenuItem.Click
        _PrintStatus = "C"
        Call Load_Gride_Production_Detailes()
        Panel3.Visible = True
        Panel2.Visible = False
        Panel1.Visible = False
    End Sub

    Private Sub UltraButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton3.Click
        _PrintStatus = "D"
        Call Load_Gride_Production_Detailes()
        Call Load_Data_Grid3()
        Panel3.Visible = False
    End Sub

    Private Sub PrintToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrintToolStripMenuItem.Click
        Dim B As New ReportDocument
        Dim A As String
        Dim StrFromDate As String
        Dim StrToDate As String
        Try
            StrFromDate = "(" & Year(_From) & ", " & VB6.Format(Month(_From), "0#") & ", " & VB6.Format(CDate(_From).Day, "0#") & ", 00, 00, 00)"
            StrToDate = "(" & Year(_to) & ", " & VB6.Format(Month(_to), "0#") & ", " & VB6.Format(CDate(_to).Day, "0#") & ", 00, 00, 00)"

            If _PrintStatus = "A" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\unpacking.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                B.SetParameterValue("To", _to)
                B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{T05Unpacking_Header.T05Date} in DateTime " & StrFromDate & " to DateTime " & StrToDate & " and {T05Unpacking_Header.T05Status}='A' "
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _PrintStatus = "B" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\unpacking.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                B.SetParameterValue("To", _to)
                B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{T05Unpacking_Header.T05Date} in DateTime " & StrFromDate & " to DateTime " & StrToDate & " and {T05Unpacking_Header.T05Status}='A' and {View_Production_Items.M14Item_Code} ='" & _Itemcode & "' "
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _PrintStatus = "C" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\unpacking1.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                B.SetParameterValue("To", _to)
                B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{T05Unpacking_Header.T05Date} in DateTime " & StrFromDate & " to DateTime " & StrToDate & " and {T05Unpacking_Header.T05Status}='A' "
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
End Class