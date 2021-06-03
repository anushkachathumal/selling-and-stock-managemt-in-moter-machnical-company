Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports CrystalDecisions.CrystalReports.Engine
Public Class frmrptStock_IN
    Dim c_dataCustomer1 As DataTable
    Dim c_dataCustomer2 As DataTable
    Dim _PrintStatus As String
    Dim _From As Date
    Dim _To As Date
    Dim _Itemcode As String

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
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
            Sql = "select * from View_Production_Items where M14Status='A' and M14Item_Name='" & cboItem.Text & "'  "
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

    Function Load_Gride_StockIN()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTable_rptStockIN
        UltraGrid1.DataSource = c_dataCustomer1
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 90
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(3).Width = 160
            .DisplayLayout.Bands(0).Columns(3).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(4).Width = 70
            .DisplayLayout.Bands(0).Columns(4).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(5).Width = 210
            .DisplayLayout.Bands(0).Columns(5).AutoEdit = False

            '.DisplayLayout.Bands(0).Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
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


    Function Load_Gride_StockIN_Summery()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTable_rptStockIN_Summery
        UltraGrid1.DataSource = c_dataCustomer1
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 160
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 90
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
           

            '.DisplayLayout.Bands(0).Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
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


    Private Sub frmrptStock_IN_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_Gride_StockIN()
        txtFrom.Text = Today
        txtTo.Text = Today
        txtDate3.Text = Today
        txtDate4.Text = Today
        Call Load_Combo()
    End Sub

    Private Sub DetailesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DetailesToolStripMenuItem.Click
        Call Load_Gride_StockIN()
        _PrintStatus = "A"
        Panel1.Visible = True
        Panel2.Visible = False
    End Sub

    Private Sub UltraButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton3.Click
        If _PrintStatus = "A" Then
            Call Load_Gride_StockIN()
            Call Load_Data_Detailes()
            _From = txtFrom.Text
            _To = txtTo.Text
            Panel1.Visible = False
        ElseIf _PrintStatus = "B" Then
            Call Load_Gride_StockIN_Summery()
            Call Load_Data_Detailes_Summery()
            _From = txtFrom.Text
            _To = txtTo.Text
            Panel1.Visible = False
        End If
    End Sub
    Function Load_Data_Item()
        Dim Sql As String

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
        Dim _Qty As Integer
        Dim Value As Double
        Dim _Rowcount As Integer

        Try
            Sql = "select *  from T03Stock_In inner join M14Product_Item on t03Item_Code=M14Item_Code where t03Date between '" & txtDate3.Text & "' and '" & txtDate4.Text & "' and T03Item_Code='" & _Itemcode & "' order by t03ref_no"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Qty = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                newRow("Batch No") = M01.Tables(0).Rows(i)("T03Ref_No")
                newRow("Date") = Month(M01.Tables(0).Rows(i)("T03Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T03Date")) & "/" & Year(M01.Tables(0).Rows(i)("T03Date"))
                newRow("Item Code") = M01.Tables(0).Rows(i)("T03Item_Code")
                newRow("Item Name") = M01.Tables(0).Rows(i)("M14Item_Name")
                If strUGroup = "ADMIN" Or strDisname = "MD" Or strDisname = "STOREKEEPER" Then
                    newRow("Qty") = M01.Tables(0).Rows(i)("T03Qty")
                    _Qty = _Qty + M01.Tables(0).Rows(i)("T03Qty")
                Else
                    newRow("Qty") = M01.Tables(0).Rows(i)("T03FG_Qty")
                    _Qty = _Qty + M01.Tables(0).Rows(i)("T03FG_Qty")
                End If
                newRow("Remark") = M01.Tables(0).Rows(i)("T03Remark")
                c_dataCustomer1.Rows.Add(newRow)

                i = i + 1
            Next
            Dim newRow1 As DataRow = c_dataCustomer1.NewRow
            newRow1("Batch No") = ""
            c_dataCustomer1.Rows.Add(newRow1)

            Dim newRow2 As DataRow = c_dataCustomer1.NewRow
            newRow2("Batch No") = "Total"
            newRow2("Qty") = _Qty

            c_dataCustomer1.Rows.Add(newRow2)

            _Rowcount = UltraGrid1.Rows.Count
            con.close()
            UltraGrid1.Rows(_Rowcount - 1).Cells(0).Appearance.BackColor = Color.Blue
            UltraGrid1.Rows(_Rowcount - 1).Cells(1).Appearance.BackColor = Color.Blue
            UltraGrid1.Rows(_Rowcount - 1).Cells(2).Appearance.BackColor = Color.Blue
            UltraGrid1.Rows(_Rowcount - 1).Cells(3).Appearance.BackColor = Color.Blue
            UltraGrid1.Rows(_Rowcount - 1).Cells(4).Appearance.BackColor = Color.Blue
            UltraGrid1.Rows(_Rowcount - 1).Cells(5).Appearance.BackColor = Color.Blue
            UltraGrid1.Rows(_Rowcount - 1).Cells(0).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            UltraGrid1.Rows(_Rowcount - 1).Cells(1).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            UltraGrid1.Rows(_Rowcount - 1).Cells(2).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            UltraGrid1.Rows(_Rowcount - 1).Cells(3).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            UltraGrid1.Rows(_Rowcount - 1).Cells(4).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            UltraGrid1.Rows(_Rowcount - 1).Cells(5).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Function Load_Data_Detailes()
        Dim Sql As String

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
        Dim _Qty As Integer
        Dim Value As Double
        Dim _Rowcount As Integer

        Try
            Sql = "select *  from T03Stock_In inner join M14Product_Item on t03Item_Code=M14Item_Code where t03Date between '" & txtFrom.Text & "' and '" & txtTo.Text & "' order by t03ref_no"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Qty = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                newRow("Batch No") = M01.Tables(0).Rows(i)("T03Ref_No")
                newRow("Date") = Month(M01.Tables(0).Rows(i)("T03Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T03Date")) & "/" & Year(M01.Tables(0).Rows(i)("T03Date"))
                newRow("Item Code") = M01.Tables(0).Rows(i)("T03Item_Code")
                newRow("Item Name") = M01.Tables(0).Rows(i)("M14Item_Name")
                If strUGroup = "ADMIN" Or strDisname = "MD" Or strDisname = "STOREKEEPER" Then
                    newRow("Qty") = M01.Tables(0).Rows(i)("T03Qty")
                    _Qty = _Qty + M01.Tables(0).Rows(i)("T03Qty")
                Else
                    newRow("Qty") = M01.Tables(0).Rows(i)("T03FG_Qty")
                    _Qty = _Qty + M01.Tables(0).Rows(i)("T03FG_Qty")
                End If
                newRow("Remark") = M01.Tables(0).Rows(i)("T03Remark")
                c_dataCustomer1.Rows.Add(newRow)

                i = i + 1
            Next
            Dim newRow1 As DataRow = c_dataCustomer1.NewRow
            newRow1("Batch No") = ""
            c_dataCustomer1.Rows.Add(newRow1)

            Dim newRow2 As DataRow = c_dataCustomer1.NewRow
            newRow2("Batch No") = "Total"
            newRow2("Qty") = _Qty

            c_dataCustomer1.Rows.Add(newRow2)

            _Rowcount = UltraGrid1.Rows.Count
            con.close()
            UltraGrid1.Rows(_Rowcount - 1).Cells(0).Appearance.BackColor = Color.Blue
            UltraGrid1.Rows(_Rowcount - 1).Cells(1).Appearance.BackColor = Color.Blue
            UltraGrid1.Rows(_Rowcount - 1).Cells(2).Appearance.BackColor = Color.Blue
            UltraGrid1.Rows(_Rowcount - 1).Cells(3).Appearance.BackColor = Color.Blue
            UltraGrid1.Rows(_Rowcount - 1).Cells(4).Appearance.BackColor = Color.Blue
            UltraGrid1.Rows(_Rowcount - 1).Cells(5).Appearance.BackColor = Color.Blue
            UltraGrid1.Rows(_Rowcount - 1).Cells(0).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            UltraGrid1.Rows(_Rowcount - 1).Cells(1).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            UltraGrid1.Rows(_Rowcount - 1).Cells(2).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            UltraGrid1.Rows(_Rowcount - 1).Cells(3).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            UltraGrid1.Rows(_Rowcount - 1).Cells(4).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            UltraGrid1.Rows(_Rowcount - 1).Cells(5).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Function Load_Data_Detailes_CurrentMonth()
        Dim Sql As String

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
        Dim _Qty As Integer
        Dim Value As Double
        Dim _Rowcount As Integer

        Try
            Sql = "select *  from T03Stock_In inner join M14Product_Item on t03Item_Code=M14Item_Code where month(t03Date)='" & Month(Today) & "' and year(t03date)='" & Year(Today) & "' order by t03ref_no"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Qty = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                newRow("Batch No") = M01.Tables(0).Rows(i)("T03Ref_No")
                newRow("Date") = Month(M01.Tables(0).Rows(i)("T03Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T03Date")) & "/" & Year(M01.Tables(0).Rows(i)("T03Date"))
                newRow("Item Code") = M01.Tables(0).Rows(i)("T03Item_Code")
                newRow("Item Name") = M01.Tables(0).Rows(i)("M14Item_Name")
                If strUGroup = "ADMIN" Or strDisname = "MD" Or strDisname = "STOREKEEPER" Then
                    newRow("Qty") = M01.Tables(0).Rows(i)("T03Qty")
                    _Qty = _Qty + M01.Tables(0).Rows(i)("T03Qty")
                Else
                    newRow("Qty") = M01.Tables(0).Rows(i)("T03FG_Qty")
                    _Qty = _Qty + M01.Tables(0).Rows(i)("T03FG_Qty")
                End If
                newRow("Remark") = M01.Tables(0).Rows(i)("T03Remark")
                c_dataCustomer1.Rows.Add(newRow)

                i = i + 1
            Next
            Dim newRow1 As DataRow = c_dataCustomer1.NewRow
            newRow1("Batch No") = ""
            c_dataCustomer1.Rows.Add(newRow1)

            Dim newRow2 As DataRow = c_dataCustomer1.NewRow
            newRow2("Batch No") = "Total"
            newRow2("Qty") = _Qty

            c_dataCustomer1.Rows.Add(newRow2)

            _Rowcount = UltraGrid1.Rows.Count
            con.close()
            UltraGrid1.Rows(_Rowcount - 1).Cells(0).Appearance.BackColor = Color.Blue
            UltraGrid1.Rows(_Rowcount - 1).Cells(1).Appearance.BackColor = Color.Blue
            UltraGrid1.Rows(_Rowcount - 1).Cells(2).Appearance.BackColor = Color.Blue
            UltraGrid1.Rows(_Rowcount - 1).Cells(3).Appearance.BackColor = Color.Blue
            UltraGrid1.Rows(_Rowcount - 1).Cells(4).Appearance.BackColor = Color.Blue
            UltraGrid1.Rows(_Rowcount - 1).Cells(5).Appearance.BackColor = Color.Blue
            UltraGrid1.Rows(_Rowcount - 1).Cells(0).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            UltraGrid1.Rows(_Rowcount - 1).Cells(1).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            UltraGrid1.Rows(_Rowcount - 1).Cells(2).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            UltraGrid1.Rows(_Rowcount - 1).Cells(3).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            UltraGrid1.Rows(_Rowcount - 1).Cells(4).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            UltraGrid1.Rows(_Rowcount - 1).Cells(5).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Function Load_Data_Detailes_Summery()
        Dim Sql As String

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
        Dim _Qty As Integer
        Dim Value As Double
        Dim _Rowcount As Integer

        Try
            Sql = "select T03Item_Code,max(M14Item_Name) as M14Item_Name,sum(T03Qty) as T03Qty,sum(T03FG_Qty) as Qty1  from T03Stock_In inner join M14Product_Item on t03Item_Code=M14Item_Code where t03Date between '" & txtFrom.Text & "' and '" & txtTo.Text & "' group by T03Item_Code"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Qty = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow
                newRow("Item Code") = M01.Tables(0).Rows(i)("T03Item_Code")
                newRow("Item Name") = M01.Tables(0).Rows(i)("M14Item_Name")
                If strUGroup = "ADMIN" Or strDisname = "MD" Or strDisname = "STOREKEEPER" Then
                    newRow("Qty") = M01.Tables(0).Rows(i)("T03Qty")
                    _Qty = _Qty + M01.Tables(0).Rows(i)("T03Qty")
                Else
                    newRow("Qty") = M01.Tables(0).Rows(i)("Qty1")
                    _Qty = _Qty + M01.Tables(0).Rows(i)("Qty1")
                End If
                c_dataCustomer1.Rows.Add(newRow)

                i = i + 1
            Next
            Dim newRow1 As DataRow = c_dataCustomer1.NewRow
            newRow1("Item Code") = ""
            c_dataCustomer1.Rows.Add(newRow1)

            Dim newRow2 As DataRow = c_dataCustomer1.NewRow
            newRow2("Item Code") = "Total"
            newRow2("Qty") = _Qty

            c_dataCustomer1.Rows.Add(newRow2)

            _Rowcount = UltraGrid1.Rows.Count
            con.close()
            UltraGrid1.Rows(_Rowcount - 1).Cells(0).Appearance.BackColor = Color.Blue
            UltraGrid1.Rows(_Rowcount - 1).Cells(1).Appearance.BackColor = Color.Blue
            UltraGrid1.Rows(_Rowcount - 1).Cells(2).Appearance.BackColor = Color.Blue
            ' UltraGrid1.Rows(_Rowcount - 1).Cells(3).Appearance.BackColor = Color.Blue
            'UltraGrid1.Rows(_Rowcount - 1).Cells(4).Appearance.BackColor = Color.Blue
            'UltraGrid1.Rows(_Rowcount - 1).Cells(5).Appearance.BackColor = Color.Blue
            UltraGrid1.Rows(_Rowcount - 1).Cells(0).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            UltraGrid1.Rows(_Rowcount - 1).Cells(1).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            UltraGrid1.Rows(_Rowcount - 1).Cells(2).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            'UltraGrid1.Rows(_Rowcount - 1).Cells(3).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            'UltraGrid1.Rows(_Rowcount - 1).Cells(4).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            'UltraGrid1.Rows(_Rowcount - 1).Cells(5).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Private Sub RefreshToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RefreshToolStripMenuItem.Click
        Call Load_Gride_StockIN()
        Panel1.Visible = False
        Panel2.Visible = False
        _Itemcode = ""
    End Sub

    Private Sub SummeryToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SummeryToolStripMenuItem.Click
        Call Load_Gride_StockIN()
        _PrintStatus = "B"
        Panel1.Visible = True
        Panel2.Visible = False
    End Sub

    Private Sub AccordingToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AccordingToolStripMenuItem.Click
        Call Load_Gride_StockIN()
        _PrintStatus = "C"
        Call Load_Data_Detailes_CurrentMonth()
        Panel1.Visible = False
        Panel2.Visible = False

        ' Panel1.Visible = True
    End Sub

    Private Sub UsingItemToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UsingItemToolStripMenuItem.Click
        Call Load_Gride_StockIN()
        cboItem.Text = ""
        Panel1.Visible = False
        Panel2.Visible = True
    End Sub


    Function Load_Combo()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M14Item_Name as [##] from View_Production_Items where M14Status='A' and Category='PI' order by M14Item_Code "
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

    Private Sub UltraButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton1.Click
        If Search_Itemcode() = True Then
            Call Load_Gride_StockIN()
            Call Load_Data_Item()
            Panel2.Visible = False
            _PrintStatus = "D"

            _From = txtDate3.Text
            _To = txtDate4.Text
        Else
            MsgBox("Please select item name", MsgBoxStyle.Information, "Information .....")
        End If
    End Sub

    Private Sub PrintToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrintToolStripMenuItem.Click
        Dim B As New ReportDocument
        Dim A As String
        Dim StrFromDate As String
        Dim StrToDate As String

        Try
            StrFromDate = "(" & Year(_From) & ", " & VB6.Format(Month(_From), "0#") & ", " & VB6.Format(CDate(_From).Day, "0#") & ", 00, 00, 00)"
            StrToDate = "(" & Year(_To) & ", " & VB6.Format(Month(_To), "0#") & ", " & VB6.Format(CDate(_To).Day, "0#") & ", 00, 00, 00)"

            If strUGroup = "ADMIN" Or strDisname = "MD" Or strDisname = "STOREKEEPER" Then
                If _PrintStatus = "A" Then
                    A = ConfigurationManager.AppSettings("ReportPath") + "\rptStockIN.rpt"
                    B.Load(A.ToString)
                    B.SetDatabaseLogon("sa", "tommya")
                    B.SetParameterValue("To", _To)
                    B.SetParameterValue("From", _From)
                    '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                    frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                    frmReport.CrystalReportViewer1.DisplayToolbar = True
                    frmReport.CrystalReportViewer1.SelectionFormula = "{T03Stock_In.T03Date} in DateTime " & StrFromDate & " to DateTime " & StrToDate & "  "
                    frmReport.Refresh()
                    ' frmReport.CrystalReportViewer1.PrintReport()
                    ' B.PrintToPrinter(1, True, 0, 0)
                    frmReport.MdiParent = MDIMain
                    frmReport.Show()
                ElseIf _PrintStatus = "B" Then
                    A = ConfigurationManager.AppSettings("ReportPath") + "\rptStockIN_Summery.rpt"
                    B.Load(A.ToString)
                    B.SetDatabaseLogon("sa", "tommya")
                    B.SetParameterValue("To", _To)
                    B.SetParameterValue("From", _From)
                    '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                    frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                    frmReport.CrystalReportViewer1.DisplayToolbar = True
                    frmReport.CrystalReportViewer1.SelectionFormula = "{T03Stock_In.T03Date} in DateTime " & StrFromDate & " to DateTime " & StrToDate & "  "
                    frmReport.Refresh()
                    ' frmReport.CrystalReportViewer1.PrintReport()
                    ' B.PrintToPrinter(1, True, 0, 0)
                    frmReport.MdiParent = MDIMain
                    frmReport.Show()

                ElseIf _PrintStatus = "C" Then
                    Dim L_DateofMonth As Integer
                    _From = Month(Today) & "/1/" & Year(Today)
                    StrFromDate = "(" & Year(_From) & ", " & VB6.Format(Month(_From), "0#") & ", " & VB6.Format(CDate(_From).Day, "0#") & ", 00, 00, 00)"
                    Dim EndDate As DateTime = _From.AddDays(DateTime.DaysInMonth(_From.Year, _From.Month) - 1)
                    '  MsgBox(MonthName(T01.Tables(0).Rows(i)("T01month"), True))
                    L_DateofMonth = Microsoft.VisualBasic.Day(EndDate)

                    _To = Month(_From) & "/" & L_DateofMonth & "/" & Year(Today)
                    StrToDate = "(" & Year(_To) & ", " & VB6.Format(Month(_To), "0#") & ", " & VB6.Format(CDate(_To).Day, "0#") & ", 00, 00, 00)"


                    A = ConfigurationManager.AppSettings("ReportPath") + "\rptStockIN.rpt"
                    B.Load(A.ToString)
                    B.SetDatabaseLogon("sa", "tommya")
                    B.SetParameterValue("To", _To)
                    B.SetParameterValue("From", _From)
                    '  frmReport.CayrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                    frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                    frmReport.CrystalReportViewer1.DisplayToolbar = True
                    frmReport.CrystalReportViewer1.SelectionFormula = "{T03Stock_In.T03Date} in DateTime " & StrFromDate & " to DateTime " & StrToDate & "  "
                    frmReport.Refresh()
                    ' frmReport.CrystalReportViewer1.PrintReport()
                    ' B.PrintToPrinter(1, True, 0, 0)
                    frmReport.MdiParent = MDIMain
                    frmReport.Show()
                ElseIf _PrintStatus = "D" Then
                    A = ConfigurationManager.AppSettings("ReportPath") + "\rptStockIN_1.rpt"
                    B.Load(A.ToString)
                    B.SetDatabaseLogon("sa", "tommya")
                    B.SetParameterValue("To", _To)
                    B.SetParameterValue("From", _From)
                    '  frmReport.CayrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                    frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                    frmReport.CrystalReportViewer1.DisplayToolbar = True
                    frmReport.CrystalReportViewer1.SelectionFormula = "{T03Stock_In.T03Date} in DateTime " & StrFromDate & " to DateTime " & StrToDate & " and {M14Product_Item.M14Item_Code}='" & _Itemcode & "'  "
                    frmReport.Refresh()
                    ' frmReport.CrystalReportViewer1.PrintReport()
                    ' B.PrintToPrinter(1, True, 0, 0)
                    frmReport.MdiParent = MDIMain
                    frmReport.Show()
                End If
            Else
                If _PrintStatus = "A" Then
                    A = ConfigurationManager.AppSettings("ReportPath") + "\rptStockIN_Fg.rpt"
                    B.Load(A.ToString)
                    B.SetDatabaseLogon("sa", "tommya")
                    B.SetParameterValue("To", _To)
                    B.SetParameterValue("From", _From)
                    '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                    frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                    frmReport.CrystalReportViewer1.DisplayToolbar = True
                    frmReport.CrystalReportViewer1.SelectionFormula = "{T03Stock_In.T03Date} in DateTime " & StrFromDate & " to DateTime " & StrToDate & "  "
                    frmReport.Refresh()
                    ' frmReport.CrystalReportViewer1.PrintReport()
                    ' B.PrintToPrinter(1, True, 0, 0)
                    frmReport.MdiParent = MDIMain
                    frmReport.Show()
                ElseIf _PrintStatus = "B" Then
                    A = ConfigurationManager.AppSettings("ReportPath") + "\rptStockIN_SummeryFG.rpt"
                    B.Load(A.ToString)
                    B.SetDatabaseLogon("sa", "tommya")
                    B.SetParameterValue("To", _To)
                    B.SetParameterValue("From", _From)
                    '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                    frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                    frmReport.CrystalReportViewer1.DisplayToolbar = True
                    frmReport.CrystalReportViewer1.SelectionFormula = "{T03Stock_In.T03Date} in DateTime " & StrFromDate & " to DateTime " & StrToDate & "  "
                    frmReport.Refresh()
                    ' frmReport.CrystalReportViewer1.PrintReport()
                    ' B.PrintToPrinter(1, True, 0, 0)
                    frmReport.MdiParent = MDIMain
                    frmReport.Show()

                ElseIf _PrintStatus = "C" Then
                    Dim L_DateofMonth As Integer
                    _From = Month(Today) & "/1/" & Year(Today)
                    StrFromDate = "(" & Year(_From) & ", " & VB6.Format(Month(_From), "0#") & ", " & VB6.Format(CDate(_From).Day, "0#") & ", 00, 00, 00)"
                    Dim EndDate As DateTime = _From.AddDays(DateTime.DaysInMonth(_From.Year, _From.Month) - 1)
                    '  MsgBox(MonthName(T01.Tables(0).Rows(i)("T01month"), True))
                    L_DateofMonth = Microsoft.VisualBasic.Day(EndDate)

                    _To = Month(_From) & "/" & L_DateofMonth & "/" & Year(Today)
                    StrToDate = "(" & Year(_To) & ", " & VB6.Format(Month(_To), "0#") & ", " & VB6.Format(CDate(_To).Day, "0#") & ", 00, 00, 00)"


                    A = ConfigurationManager.AppSettings("ReportPath") + "\rptStockIN_Fg.rpt"
                    B.Load(A.ToString)
                    B.SetDatabaseLogon("sa", "tommya")
                    B.SetParameterValue("To", _To)
                    B.SetParameterValue("From", _From)
                    '  frmReport.CayrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                    frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                    frmReport.CrystalReportViewer1.DisplayToolbar = True
                    frmReport.CrystalReportViewer1.SelectionFormula = "{T03Stock_In.T03Date} in DateTime " & StrFromDate & " to DateTime " & StrToDate & "  "
                    frmReport.Refresh()
                    ' frmReport.CrystalReportViewer1.PrintReport()
                    ' B.PrintToPrinter(1, True, 0, 0)
                    frmReport.MdiParent = MDIMain
                    frmReport.Show()
                ElseIf _PrintStatus = "D" Then
                    A = ConfigurationManager.AppSettings("ReportPath") + "\rptStockIN_Fg1.rpt"
                    B.Load(A.ToString)
                    B.SetDatabaseLogon("sa", "tommya")
                    B.SetParameterValue("To", _To)
                    B.SetParameterValue("From", _From)
                    '  frmReport.CayrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                    frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                    frmReport.CrystalReportViewer1.DisplayToolbar = True
                    frmReport.CrystalReportViewer1.SelectionFormula = "{T03Stock_In.T03Date} in DateTime " & StrFromDate & " to DateTime " & StrToDate & " and {M14Product_Item.M14Item_Code}='" & _Itemcode & "'  "
                    frmReport.Refresh()
                    ' frmReport.CrystalReportViewer1.PrintReport()
                    ' B.PrintToPrinter(1, True, 0, 0)
                    frmReport.MdiParent = MDIMain
                    frmReport.Show()
                End If
            End If

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                ' con.close()
            End If
        End Try
    End Sub
End Class