Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports CrystalDecisions.CrystalReports.Engine
Public Class frmWastege_Conf
    Dim c_dataCustomer1 As DataTable
    Dim c_dataCustomer2 As DataTable
    Dim c_dataCustomer3 As DataTable
    Dim _PrintStatus As String
    Dim _MainStatus As String
    Dim _From As Date
    Dim _To As Date
    Dim _Itemcode As String
    Dim _Supplier As String
    Dim _Category As String
    Dim _Comcode As String


    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub


    Function Load_Gride()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTable_Wastage_App
        UltraGrid2.DataSource = c_dataCustomer1
        With UltraGrid2
            .DisplayLayout.Bands(0).Columns(0).Width = 70
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 90
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(3).Width = 90
            .DisplayLayout.Bands(0).Columns(3).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(4).Width = 90
            .DisplayLayout.Bands(0).Columns(4).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(8).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(9).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

        End With
    End Function

    Private Sub frmWastege_Conf_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        _Comcode = ConfigurationManager.AppSettings("LOCCode")
        Call Load_Gride()
        Call Load_Data()
        txtDate.ReadOnly = True
        txtEntry.ReadOnly = True
        txtEntry.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtDate.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtRemark.ReadOnly = True
        txtNett.ReadOnly = True
        txtNett.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
        Call Load_Gride_Item()
    End Sub

    Function Load_Data()
        Dim Sql As String

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
        Dim _Qty As Integer
        Dim Value As Double
        Dim _Rowcount As Integer
        Dim _St As String
        Dim _Total As Double
        Dim _LastRow As Integer

        Try
            Me.Cursor = Cursors.WaitCursor
            Call Load_Gride()
            Sql = "select * from T01Transaction_Header where T01Trans_Type='WT' and T01Status='A' and T01Com_Code='" & _Comcode & "' order by T01Ref_No"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Total = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow
                newRow("##") = "PENDING"
                newRow("Wastage No") = M01.Tables(0).Rows(i)("T01Grn_No")
                newRow("Date") = Month(M01.Tables(0).Rows(i)("T01Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T01Date")) & "/" & Year(M01.Tables(0).Rows(i)("T01Date"))
                'newRow("Com Invoice") = M01.Tables(0).Rows(i)("T01Invoice_No")
               
                Value = M01.Tables(0).Rows(i)("T01Net_Amount")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Net Amount") = _St
                newRow("User") = M01.Tables(0).Rows(i)("T01User")

                

                c_dataCustomer1.Rows.Add(newRow)

                i = i + 1
            Next

            con.close()
            Me.Cursor = Cursors.Arrow
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                Me.Cursor = Cursors.Arrow
                con.close()
            End If
        End Try
    End Function

    Function Load_Data_a1()
        Dim Sql As String

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
        Dim _Qty As Integer
        Dim Value As Double
        Dim _Rowcount As Integer
        Dim _St As String
        Dim _Total As Double
        Dim _LastRow As Integer
        Dim M02 As DataSet

        Try
            Me.Cursor = Cursors.WaitCursor
            Call Load_Gride()

            Sql = "SELECT CONVERT(date, tmptime) as tm,tmpRef_No FROM tmpTransaction_Log where tmp_TR='WT' AND tmpProcess='APPROVED' AND CONVERT(date, tmptime) BETWEEN '" & _From & "' AND '" & _To & "'"
            M02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            For Each DTRow1 As DataRow In M02.Tables(0).Rows
                Sql = "select * from T01Transaction_Header where T01Trans_Type='WT' and T01Status='APP' AND T01Grn_No='" & M02.Tables(0).Rows(i)("tmpRef_No") & "' order by T01Ref_No"
                M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)

                _Total = 0
                If isValidDataset(M01) Then
                    Dim newRow As DataRow = c_dataCustomer1.NewRow
                    newRow("##") = "APPROVED"
                    newRow("Wastage No") = M01.Tables(0).Rows(0)("T01Grn_No")
                    newRow("Date") = Month(M02.Tables(0).Rows(i)("TM")) & "/" & Microsoft.VisualBasic.Day(M02.Tables(0).Rows(i)("TM")) & "/" & Year(M02.Tables(0).Rows(i)("TM"))
                    'newRow("Com Invoice") = M01.Tables(0).Rows(i)("T01Invoice_No")

                    Value = M01.Tables(0).Rows(0)("T01Net_Amount")
                    _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                    newRow("Net Amount") = _St
                    newRow("User") = M01.Tables(0).Rows(0)("T01User")

                    c_dataCustomer1.Rows.Add(newRow)

                End If
              
                i = i + 1
            Next

                con.close()
                Me.Cursor = Cursors.Arrow
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                Me.Cursor = Cursors.Arrow
                con.close()
            End If
        End Try
    End Function

    Function Load_Data_a2()
        Dim Sql As String

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
        Dim _Qty As Integer
        Dim Value As Double
        Dim _Rowcount As Integer
        Dim _St As String
        Dim _Total As Double
        Dim _LastRow As Integer
        Dim M02 As DataSet

        Try
            Me.Cursor = Cursors.WaitCursor
            Call Load_Gride()

            Sql = "SELECT CONVERT(date, tmptime) as tm,tmpRef_No FROM tmpTransaction_Log where tmp_TR='WT' AND tmpProcess='REJECT' AND CONVERT(date, tmptime) BETWEEN '" & _From & "' AND '" & _To & "'"
            M02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            For Each DTRow1 As DataRow In M02.Tables(0).Rows
                Sql = "select * from T01Transaction_Header where T01Trans_Type='WT' and T01Status='I' AND T01Grn_No='" & M02.Tables(0).Rows(i)("tmpRef_No") & "' order by T01Ref_No"
                M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)

                _Total = 0
                If isValidDataset(M01) Then
                    Dim newRow As DataRow = c_dataCustomer1.NewRow
                    newRow("##") = "REJECT"
                    newRow("Wastage No") = M01.Tables(0).Rows(0)("T01Grn_No")
                    newRow("Date") = Month(M02.Tables(0).Rows(i)("TM")) & "/" & Microsoft.VisualBasic.Day(M02.Tables(0).Rows(i)("TM")) & "/" & Year(M02.Tables(0).Rows(i)("TM"))
                    'newRow("Com Invoice") = M01.Tables(0).Rows(i)("T01Invoice_No")

                    Value = M01.Tables(0).Rows(0)("T01Net_Amount")
                    _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                    newRow("Net Amount") = _St
                    newRow("User") = M01.Tables(0).Rows(i)("T01User")

                    c_dataCustomer1.Rows.Add(newRow)

                End If

                i = i + 1
            Next

            con.close()
            Me.Cursor = Cursors.Arrow
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                Me.Cursor = Cursors.Arrow
                con.close()
            End If
        End Try
    End Function

    Private Sub ResetToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ResetToolStripMenuItem.Click
        Call Load_Gride()
        Call Load_Data()
        _PrintStatus = ""
        OPR0.Visible = False
        Panel4.Visible = False
    End Sub

    Private Sub UltraGrid2_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles UltraGrid2.DoubleClickRow
        On Error Resume Next
        Dim _RowIndex As Integer
        _RowIndex = UltraGrid2.ActiveRow.Index
        OPR0.Visible = True
        txtEntry.Text = UltraGrid2.Rows(_RowIndex).Cells(2).Text
        Call Load_Gride_Item()
        Call Search_RecordsUsing_Entry()

    End Sub

    Function Search_RecordsUsing_Entry() As Boolean
        Dim result1 As String
        Dim Value As Double
        Dim SQL As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim T01 As DataSet
        Dim _St As String
        Dim I As Integer
        Dim _RefNo As Double

        Try
            SQL = "select * from View_Wastage_Header where T01Grn_No='" & txtEntry.Text & "' and T01Com_Code='" & _Comcode & "' "
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(T01) Then
                Search_RecordsUsing_Entry = True
                txtRemark.Text = T01.Tables(0).Rows(0)("T01Remark")
                txtDate.Text = T01.Tables(0).Rows(0)("T01Date")
                _RefNo = T01.Tables(0).Rows(0)("T01Ref_no")
                Value = T01.Tables(0).Rows(0)("T01Net_amount")
                txtNett.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtNett.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            End If

            SQL = "select * from View_T02Transaction where T02Ref_No=" & _RefNo & ""
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            I = 0
            For Each DTRow2 As DataRow In T01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer2.NewRow
                newRow("Item Code") = T01.Tables(0).Rows(I)("T02Item_Code")
                newRow("Item Name") = T01.Tables(0).Rows(I)("M03Item_Name")
                Value = T01.Tables(0).Rows(I)("T02cost")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Rate") = _St
                '  newRow("Retail Price") = txtSales.Text
                newRow("Qty") = T01.Tables(0).Rows(I)("T02Qty")
                Value = T01.Tables(0).Rows(I)("T02Total")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Total") = _St



                c_dataCustomer2.Rows.Add(newRow)
                I = I + 1
            Next

            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function


    Function Load_Gride_Item()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer2 = CustomerDataClass.MakeDataTable_PO
        UltraGrid3.DataSource = c_dataCustomer2
        With UltraGrid3
            .DisplayLayout.Bands(0).Columns(0).Width = 110
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 230
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False

            .DisplayLayout.Bands(0).Columns(2).Width = 70
            .DisplayLayout.Bands(0).Columns(3).Width = 70
            .DisplayLayout.Bands(0).Columns(4).Width = 70
            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            '  .DisplayLayout.Bands(0).Columns(6).Width = 90
            ' .DisplayLayout.Bands(0).Columns(7).Width = 90

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        OPR0.Visible = False
    End Sub

    Private Sub UltraGrid2_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles UltraGrid2.InitializeLayout

    End Sub

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
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
        Dim result1 As String
        Dim M01 As DataSet
        Dim Value As Double
        Dim A As String
        Dim A1 As String
        Dim B As New ReportDocument
        Dim _RefNo As Double

        Try
            Call Load_Gride_Item()
            If Search_RecordsUsing_Entry() = True Then
            Else
                MsgBox("Wrong wastage Note.Please try again", MsgBoxStyle.Information, "Information ......")
                connection.Close()
                OPR0.Visible = False
                Exit Sub
            End If
            nvcFieldList1 = "SELECT * FROM T01Transaction_Header WHERE T01Grn_No='" & txtEntry.Text & "' AND T01Com_Code='" & _Comcode & "'  AND T01Trans_Type='WT'"
            M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(M01) Then
                _RefNo = M01.Tables(0).Rows(0)("T01Ref_No")
            End If

            nvcFieldList1 = "UPDATE T01Transaction_Header SET T01STATUS='APP' WHERE T01Grn_No='" & txtEntry.Text & "' AND T01FromLoc_Code='" & _Comcode & "' AND T01Trans_Type='WT'"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)


            nvcFieldList1 = "Insert Into tmpTransaction_Log(tmp_TR,tmpProcess,tmpRef_No,tmpTime,tmpAthz,tmpUser,tmpLog)" & _
                                                              " values('WT', 'APPROVED','" & txtEntry.Text & "','" & Now & "','" & strDisname & "','" & strDisname & "','" & _Comcode & "')"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            transaction.Commit()
            A = MsgBox("Are you sure you want to print Wastage Note", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Print Dispatch Note .....")
            If A = vbYes Then
                A1 = ConfigurationManager.AppSettings("ReportPath") + "\Wastage_Note.rpt"
                B.Load(A1.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                'B.SetParameterValue("To", _To)
                'B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_Wastage_Header.T01Ref_No}  =" & _RefNo & " and {View_Wastage_Header.T01FromLoc_Code} ='" & _Comcode & "'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            End If

            connection.Close()
            OPR0.Visible = False
            Call Load_Gride()
            Call Load_Data()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
            End If
        End Try

    End Sub

    Private Sub cmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelete.Click
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
        Dim result1 As String
        Dim M01 As DataSet
        Dim Value As Double
        Dim A As String
        Dim A1 As String
        Dim B As New ReportDocument
        Dim _RefNo As Double

        Try
            If Search_RecordsUsing_Entry() = True Then
            Else
                MsgBox("Wrong wastage Note.Please try again", MsgBoxStyle.Information, "Information ......")
                connection.Close()
                OPR0.Visible = False
                Exit Sub
            End If
            nvcFieldList1 = "SELECT * FROM T01Transaction_Header WHERE T01Grn_No='" & txtEntry.Text & "' AND T01FromLoc_Code='" & _Comcode & "'  AND T01Trans_Type='WT'"
            M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(M01) Then
                _RefNo = M01.Tables(0).Rows(0)("T01Ref_No")
            End If

            A = MsgBox("Are you sure you want to reject this wastage note", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Reject ......")
            If A = vbYes Then
                nvcFieldList1 = "UPDATE T01Transaction_Header SET T01STATUS='I' WHERE T01Grn_No='" & txtEntry.Text & "' AND T01FromLoc_Code='" & _Comcode & "' AND T01Trans_Type='WT'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)


                nvcFieldList1 = "Insert Into tmpTransaction_Log(tmp_TR,tmpProcess,tmpRef_No,tmpTime,tmpAthz_Code,tmpUser,tmpLog)" & _
                                                                  " values('WT', 'REJECT','" & txtEntry.Text & "','" & Now & "','" & strDisname & "','" & strDisname & "','" & _Comcode & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                transaction.Commit()
            End If
            connection.Close()
            OPR0.Visible = False
            Call Load_Gride()
            Call Load_Data()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
            End If
        End Try

    End Sub

    Private Sub ByDateToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ByDateToolStripMenuItem.Click
        _PrintStatus = "A1"
        Panel4.Visible = True
        txtC1.Text = Today
        txtC2.Text = Today

    End Sub

    Private Sub UltraButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton4.Click
        If _PrintStatus = "A1" Then
            _From = txtC1.Text
            _To = txtC2.Text
            Panel4.Visible = False
            Call Load_Data_a1()
        ElseIf _PrintStatus = "A2" Then
            _From = txtC1.Text
            _To = txtC2.Text
            Panel4.Visible = False
            Call Load_Data_a1()
        End If
    End Sub

    Private Sub UltraButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton1.Click
        Dim A1 As String
        Dim B As New ReportDocument
        Dim _RefNo As Double
        Dim SQL As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet

        Try
            If _PrintStatus = "A1" Then

                SQL = "SELECT * FROM T01Transaction_Header WHERE T01Grn_No='" & txtEntry.Text & "' AND T01Com_Code='" & _Comcode & "'  AND T01Trans_Type='WT'"
                M01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                If isValidDataset(M01) Then
                    _RefNo = M01.Tables(0).Rows(0)("T01Ref_No")
                End If

                A1 = ConfigurationManager.AppSettings("ReportPath") + "\Wastage_Note.rpt"
                B.Load(A1.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                'B.SetParameterValue("To", _To)
                'B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_Wastage_Header.T01Ref_No}  =" & _RefNo & " and {View_Wastage_Header.T01FromLoc_Code} ='" & _Comcode & "'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()

            End If
            con.CLOSE()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.CLOSE()
            End If
        End Try
    End Sub

    Private Sub RejectToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RejectToolStripMenuItem.Click
        _PrintStatus = "A2"
        Panel4.Visible = True
        txtC1.Text = Today
        txtC2.Text = Today
    End Sub
End Class