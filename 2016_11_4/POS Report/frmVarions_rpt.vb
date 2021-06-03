Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports CrystalDecisions.CrystalReports.Engine
Public Class frmVarions_rpt
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
        c_dataCustomer1 = CustomerDataClass.MakeDataTable_Varions
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
            .DisplayLayout.Bands(0).Columns(4).Width = 90
            .DisplayLayout.Bands(0).Columns(4).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(5).Width = 70
            .DisplayLayout.Bands(0).Columns(5).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(6).Width = 90
            .DisplayLayout.Bands(0).Columns(6).AutoEdit = False
            '.DisplayLayout.Bands(0).Columns(7).Width = 70
            '.DisplayLayout.Bands(0).Columns(7).AutoEdit = False
            ''.DisplayLayout.Bands(0).Columns(8).Width = 70
            '.DisplayLayout.Bands(0).Columns(8).AutoEdit = False
            '.DisplayLayout.Bands(0).Columns(9).Width = 90
            '.DisplayLayout.Bands(0).Columns(9).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

        End With
    End Function

    Private Sub frmVarions_rpt_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        _Comcode = ConfigurationManager.AppSettings("LOCCode")
        Call Load_Gride()
    End Sub

    Function Load_Data()
        Dim Sql As String

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
        Dim _Qty As Double
        Dim Value As Double
        Dim _Rowcount As Integer
        Dim _St As String
        Dim m02 As DataSet
        Dim m03 As DataSet
        Dim _NEW As Double

        Try
            Sql = "select S01Item_Code,max(S01Date ) as D_Date,max(S01Qty) as Qty  from S01Stock_Balance where S01Trans_Type='ob' and S01Status='A' and S01Loc_Code='" & _Comcode & "' AND S01Date BETWEEN '" & txtC1.Text & "' AND '" & txtC2.Text & "' group by S01Item_Code  order by d_date"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows

                Sql = "SELECT * FROM  S01Stock_Balance INNER JOIN M03Item_Master ON S01Item_Code=M03Item_Code  WHERE S01Date BETWEEN '" & txtC1.Text & "'  AND '" & txtC2.Text & "' And S01Status='A' AND S01Item_Code='" & Trim(M01.Tables(0).Rows(i)("S01Item_Code")) & "' and S01Loc_Code='" & _Comcode & "' AND S01Trans_Type='ob'"
                m02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(m02) Then
                    Dim newRow As DataRow = c_dataCustomer1.NewRow
                    newRow("Date(L/Up)") = Month(m02.Tables(0).Rows(0)("S01Date")) & "/" & Microsoft.VisualBasic.Day(m02.Tables(0).Rows(0)("S01Date")) & "/" & Year(m02.Tables(0).Rows(0)("S01Date"))
                    newRow("Item Code") = Trim(m02.Tables(0).Rows(0)("S01Item_Code"))
                    newRow("Item Name") = Trim(m02.Tables(0).Rows(0)("M03Item_Name"))
                    Value = Trim(m02.Tables(0).Rows(0)("M03Cost_Price"))
                    _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    newRow("Cost Price") = _St
                    _Qty = 0
                    Sql = "SELECT SUM(S01Qty) as Qty FROM  S01Stock_Balance INNER JOIN M03Item_Master ON S01Item_Code=M03Item_Code  WHERE S01Date < '" & M01.Tables(0).Rows(i)("D_Date") & "' And S01Status<>'CLOSE' AND S01Item_Code='" & Trim(M01.Tables(0).Rows(i)("S01Item_Code")) & "' and S01Loc_Code='" & _Comcode & "'  GROUP BY S01Item_Code"
                    m03 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(m03) Then
                        _Qty = m03.Tables(0).Rows(0)("qTY")
                        newRow("Last Qty") = _Qty
                    End If

                    _NEW = 0
                    Sql = "SELECT SUM(S01Qty) as Qty FROM  S01Stock_Balance INNER JOIN M03Item_Master ON S01Item_Code=M03Item_Code  WHERE  S01Status='A' AND S01Item_Code='" & Trim(M01.Tables(0).Rows(i)("S01Item_Code")) & "' and S01Loc_Code='" & _Comcode & "' AND S01Trans_Type='ob' GROUP BY S01Item_Code"
                    m03 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(m03) Then
                        _NEW = m03.Tables(0).Rows(0)("qTY")
                        newRow("New Qty") = _NEW
                    End If
                    'If _Qty > 0 Then
                    newRow("Variance") = _NEW - _Qty
                    ' Else
                    ' newRow("Variance") = _NEW + _Qty
                    ' End If

                    Value = (_NEW - _Qty) * CDbl(m02.Tables(0).Rows(0)("M03Cost_Price"))
                    _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    newRow("Variance Value") = _St

                    c_dataCustomer1.Rows.Add(newRow)
                End If

                i = i + 1
            Next
            Panel4.Visible = False
            Call Save_Date()
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Function Save_Date()
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

        Try
            Cursor = Cursors.WaitCursor
            nvcFieldList1 = "DELETE FROM R06Report"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            i = 0
            For Each uRow As UltraGridRow In UltraGrid1.Rows
                nvcFieldList1 = "Insert Into R06Report(R06DATE,R06Item_Code,R06Last,R06New,R06Location)" & _
                                                                " values('" & UltraGrid1.Rows(i).Cells(0).Value & "', '" & UltraGrid1.Rows(i).Cells(1).Value & "','" & (UltraGrid1.Rows(i).Cells(4).Value) & "','" & (UltraGrid1.Rows(i).Cells(5).Value) & "','" & _Comcode & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                i = i + 1
            Next

            transaction.Commit()
            MsgBox("Report genarated successfully", MsgBoxStyle.Information, "Information .......")
            Cursor = Cursors.Arrow
            connection.Close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.Close()
            End If
        End Try
    End Function

    Private Sub UsingDateToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UsingDateToolStripMenuItem.Click
        Call Load_Gride()
        ' Call Load_Data()
        Panel4.Visible = True
        txtC1.Text = Today
        txtC2.Text = Today
    End Sub

    Private Sub UltraButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton4.Click
        Call Load_Gride()
        Call Load_Data()
        _From = txtC1.Text
        _To = txtC2.Text
    End Sub

    Private Sub ResetToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ResetToolStripMenuItem.Click
        Panel4.Visible = False
        Call Load_Gride()
    End Sub

    Private Sub PrintToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrintToolStripMenuItem.Click
        Dim B As New ReportDocument
        Dim A As String
        Dim StrFromDate As String
        Dim StrToDate As String
        Try
            StrFromDate = "(" & Year(_From) & ", " & VB6.Format(Month(_From), "0#") & ", " & VB6.Format(CDate(_From).Day, "0#") & ", 00, 00, 00)"
            StrToDate = "(" & Year(_To) & ", " & VB6.Format(Month(_To), "0#") & ", " & VB6.Format(CDate(_To).Day, "0#") & ", 00, 00, 00)"


            A = ConfigurationManager.AppSettings("ReportPath") + "\Variance.rpt"
            B.Load(A.ToString)
            B.SetDatabaseLogon("sa", "tommya")
            B.SetParameterValue("To", _To)
            B.SetParameterValue("From", _From)
            '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
            frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
            frmReport.CrystalReportViewer1.DisplayToolbar = True
            frmReport.CrystalReportViewer1.SelectionFormula = "{R06Report.R06Location}='" & _Comcode & "'"
            frmReport.Refresh()
            ' frmReport.CrystalReportViewer1.PrintReport()
            ' B.PrintToPrinter(1, True, 0, 0)
            frmReport.MdiParent = MDIMain
            frmReport.Show()



        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                ' con.close()
            End If
        End Try
    End Sub


End Class