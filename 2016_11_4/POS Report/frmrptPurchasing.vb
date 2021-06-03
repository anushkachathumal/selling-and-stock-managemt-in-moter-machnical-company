Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports CrystalDecisions.CrystalReports.Engine
Public Class frmrptPurchasing
    Dim c_dataCustomer1 As DataTable
    Dim c_dataCustomer2 As DataTable
    Dim _PrintStatus As String
    Dim _From As Date
    Dim _To As Date
    Dim _Itemcode As String
    Dim _Com_Code As String

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Function Load_Gride()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTable_Purchasing
        UltraGrid1.DataSource = c_dataCustomer1
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 70
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 90
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(3).Width = 190
            .DisplayLayout.Bands(0).Columns(3).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(4).Width = 90
            .DisplayLayout.Bands(0).Columns(4).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(5).Width = 130
            .DisplayLayout.Bands(0).Columns(5).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(6).Width = 70
            .DisplayLayout.Bands(0).Columns(6).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(7).Width = 70
            .DisplayLayout.Bands(0).Columns(7).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(8).Width = 70
            .DisplayLayout.Bands(0).Columns(8).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(9).Width = 90
            .DisplayLayout.Bands(0).Columns(9).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(8).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(9).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

        End With
    End Function
    Function Load_Item()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M03Item_Code as [##] from M03Item_Master where m03Status='A' and M03Com_Code='" & _Com_Code & "' order by M03Item_Code "
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
    Private Sub frmrptPurchasing_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        _Com_Code = ConfigurationManager.AppSettings("LOCCode")
        Call Load_Gride()
        txtDate3.Text = Today
        txtDate4.Text = Today
        txtA1.Text = Today
        txtA2.Text = Today
        Call Load_Supplier()
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
            Sql = "select M02Cat_Name as [##] from M02Category  "
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

    Function Load_Supplier()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M09Name as [##] from M09Supplier where M09Loc_Code='" & _Com_Code & "' "
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

    Private Sub UsingDateToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UsingDateToolStripMenuItem.Click
        Panel3.Visible = True
        Panel1.Visible = False
        Panel2.Visible = False
        Panel4.Visible = False
    End Sub

    Private Sub UltraButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton2.Click
        Call Load_Gride()
        _PrintStatus = "A1"
        Call Load_Data_A1()
        Panel3.Visible = False
    End Sub

    Function Load_Data_A1()
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
            Sql = "select *  from View_Purchasing where t01date between '" & txtDate3.Text & "' and '" & txtDate4.Text & "' and T01Com_Code='" & _Com_Code & "' order by T01Ref_No"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Qty = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                newRow("Ref No") = M01.Tables(0).Rows(i)("T01Ref_No")
                newRow("Date") = Month(M01.Tables(0).Rows(i)("T01Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T01Date")) & "/" & Year(M01.Tables(0).Rows(i)("T01Date"))
                newRow("Com Invoice") = M01.Tables(0).Rows(i)("T01Invoice_No")
                newRow("Supplier") = M01.Tables(0).Rows(i)("M09Name")
                newRow("Item Code") = M01.Tables(0).Rows(i)("m03Item_Code")
                newRow("Item Name") = M01.Tables(0).Rows(i)("M03Item_Name")
                Value = M01.Tables(0).Rows(i)("T02Qty")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Qty") = _St

                Value = M01.Tables(0).Rows(i)("T02Free_Issue")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Free Issue") = _St
                Value = M01.Tables(0).Rows(i)("T02Cost")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Cost Price") = _St

                Value = M01.Tables(0).Rows(i)("total")
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

    Function Load_Data_A2()
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
            Sql = "select *  from View_Purchasing where t01date between '" & txtA1.Text & "' and '" & txtA2.Text & "' and M09Name='" & Trim(cboSupplier.Text) & "' and T01Com_Code='" & _Com_Code & "' order by T01Ref_No"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Qty = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                newRow("Ref No") = M01.Tables(0).Rows(i)("T01Ref_No")
                newRow("Date") = Month(M01.Tables(0).Rows(i)("T01Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T01Date")) & "/" & Year(M01.Tables(0).Rows(i)("T01Date"))
                newRow("Com Invoice") = M01.Tables(0).Rows(i)("T01Invoice_No")
                newRow("Supplier") = M01.Tables(0).Rows(i)("M09Name")
                newRow("Item Code") = M01.Tables(0).Rows(i)("m03Item_Code")
                newRow("Item Name") = M01.Tables(0).Rows(i)("M03Item_Name")
                Value = M01.Tables(0).Rows(i)("T02Qty")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Qty") = _St

                Value = M01.Tables(0).Rows(i)("T02Free_Issue")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Free Issue") = _St
                Value = M01.Tables(0).Rows(i)("T02Cost")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Cost Price") = _St

                Value = M01.Tables(0).Rows(i)("total")
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

    Function Load_Data_A3()
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
            Sql = "select *  from View_Purchasing where t01date between '" & txtB1.Text & "' and '" & txtB2.Text & "' and M02Cat_Name='" & Trim(cboCategory.Text) & "' and T01Com_Code='" & _Com_Code & "' order by T01Ref_No"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Qty = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                newRow("Ref No") = M01.Tables(0).Rows(i)("T01Ref_No")
                newRow("Date") = Month(M01.Tables(0).Rows(i)("T01Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T01Date")) & "/" & Year(M01.Tables(0).Rows(i)("T01Date"))
                newRow("Com Invoice") = M01.Tables(0).Rows(i)("T01Invoice_No")
                newRow("Supplier") = M01.Tables(0).Rows(i)("M09Name")
                newRow("Item Code") = M01.Tables(0).Rows(i)("m03Item_Code")
                newRow("Item Name") = M01.Tables(0).Rows(i)("M03Item_Name")
                Value = M01.Tables(0).Rows(i)("T02Qty")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Qty") = _St

                Value = M01.Tables(0).Rows(i)("T02Free_Issue")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Free Issue") = _St
                Value = M01.Tables(0).Rows(i)("T02Cost")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Cost Price") = _St

                Value = M01.Tables(0).Rows(i)("total")
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

    Function Load_Data_A5(ByVal strInv0 As String)
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

            Sql = "select *  from View_Purchasing where  T01Ref_No='" & strInv0 & "' and T01Com_Code='" & _Com_Code & "' order by T01Ref_No"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Qty = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                newRow("Ref No") = M01.Tables(0).Rows(i)("T01Ref_No")
                newRow("Date") = Month(M01.Tables(0).Rows(i)("T01Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T01Date")) & "/" & Year(M01.Tables(0).Rows(i)("T01Date"))
                newRow("Com Invoice") = M01.Tables(0).Rows(i)("T01Invoice_No")
                newRow("Supplier") = M01.Tables(0).Rows(i)("M09Name")
                newRow("Item Code") = M01.Tables(0).Rows(i)("m03Item_Code")
                newRow("Item Name") = M01.Tables(0).Rows(i)("M03Item_Name")
                Value = M01.Tables(0).Rows(i)("T02Qty")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Qty") = _St

                Value = M01.Tables(0).Rows(i)("T02Free_Issue")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Free Issue") = _St
                Value = M01.Tables(0).Rows(i)("T02Cost")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Cost Price") = _St

                Value = M01.Tables(0).Rows(i)("total")
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

    Function Load_Data_A4()
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
            Sql = "select *  from View_Purchasing where t01date between '" & txtC1.Text & "' and '" & txtC2.Text & "' and M03Item_Code='" & Trim(cboItem.Text) & "' and T01Com_Code='" & _Com_Code & "' order by T01Ref_No"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Qty = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                newRow("Ref No") = M01.Tables(0).Rows(i)("T01Ref_No")
                newRow("Date") = Month(M01.Tables(0).Rows(i)("T01Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T01Date")) & "/" & Year(M01.Tables(0).Rows(i)("T01Date"))
                newRow("Com Invoice") = M01.Tables(0).Rows(i)("T01Invoice_No")
                newRow("Supplier") = M01.Tables(0).Rows(i)("M09Name")
                newRow("Item Code") = M01.Tables(0).Rows(i)("m03Item_Code")
                newRow("Item Name") = M01.Tables(0).Rows(i)("M03Item_Name")
                Value = M01.Tables(0).Rows(i)("T02Qty")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Qty") = _St

                Value = M01.Tables(0).Rows(i)("T02Free_Issue")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Free Issue") = _St
                Value = M01.Tables(0).Rows(i)("T02Cost")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Cost Price") = _St

                Value = M01.Tables(0).Rows(i)("total")
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
    Private Sub ResetToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ResetToolStripMenuItem.Click
        Call Load_Gride()
        Panel3.Visible = False
        Panel1.Visible = False
        Panel2.Visible = False
        Panel4.Visible = False
    End Sub

    Private Sub UsingSupplierToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UsingSupplierToolStripMenuItem.Click
        Panel3.Visible = False
        Panel1.Visible = True
        Panel2.Visible = False
        Panel4.Visible = False
    End Sub

    Private Sub UltraButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton1.Click
        Call Load_Gride()
        _PrintStatus = "A2"
        Call Load_Data_A2()
        Panel1.Visible = False
    End Sub

    Private Sub UsingCategoryToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UsingCategoryToolStripMenuItem.Click
        Panel3.Visible = False
        Panel1.Visible = False
        Panel2.Visible = True
        Panel4.Visible = False
    End Sub

    Private Sub UltraButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton3.Click
        Call Load_Gride()
        _PrintStatus = "A3"
        Call Load_Data_A3()
        Panel2.Visible = False
    End Sub

    Private Sub UsingItemNameToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UsingItemNameToolStripMenuItem.Click
        Panel3.Visible = False
        Panel1.Visible = False
        Panel2.Visible = False
        Panel4.Visible = True
    End Sub

    Private Sub UltraButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton4.Click
        Call Load_Gride()
        _PrintStatus = "A4"
        Call Load_Data_A4()
        Panel4.Visible = False
    End Sub

    Private Sub UltraGrid1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles UltraGrid1.DoubleClick
        On Error Resume Next
        Dim _RowIndex As Integer
        Dim strInv0 As String
        _RowIndex = UltraGrid1.ActiveRow.Index
        strInv0 = UltraGrid1.Rows(_RowIndex).Cells(0).Text
        Call Load_Gride()
        Call Load_Data_A5(strInv0)
    End Sub

    Private Sub UltraGrid1_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles UltraGrid1.InitializeLayout

    End Sub

    Private Sub PrintToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrintToolStripMenuItem.Click

    End Sub
End Class