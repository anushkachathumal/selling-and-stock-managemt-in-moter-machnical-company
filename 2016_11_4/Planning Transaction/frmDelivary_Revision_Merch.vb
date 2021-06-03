'// Reverce Delivary Quotation
'// Development Date - 09.04.2014
'// Developed by - Suranga Wijesinghe
'// Audit by     - Amila Priyankara - TJL
'// Referance Table - T02Delivary_Quat_Header
'//                 - T03Delivary_Quat_Flutter
'//                 - USERS     
'//---------------------------------------------------------->>>
'Revice Delivary quation for Merchant according to Customer request

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

Public Class frmDelivary_Revision_Merch
    Dim c_dataCustomer As DataTable
    Dim Clicked As String
    Dim _RefNo As Integer
    Dim advancedSearchTag As String = ""
    Dim strPlanner As String
    Dim _Del_ReqNo As Integer
    Dim strParameter As Integer

    Function Load_Sales_Order()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        'Load sales order to cboSO combobox

        Try
            Sql = "select T02OrderNo as [Sales Order] from T02Delivary_Quat_Header innner join T01Delivary_Request on T01RefNo=T02Del_Req_No where T02Status='A' and T01User='" & Trim(strDisname) & "' group by T02OrderNo"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboSales_Order
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 180
                ' .Rows.Band.Columns(1).Width = 260


            End With
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""


        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub frmDelivary_Revision_Merch_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_Sales_Order()
        txtDepartment.ReadOnly = True
        txtPlanner.ReadOnly = True
        txtPO.ReadOnly = True

        txtPO.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        Call Load_Gride_With_Data()

    End Sub

    Function Search_Sales_Order() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim M02 As DataSet

        'Search PO Number

        Try
            Sql = "select * from M01Sales_Order_SAP where CONVERT(INT,M01Sales_Order)='" & Trim(cboSales_Order.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                Search_Sales_Order = True
                txtPO.Text = M01.Tables(0).Rows(0)("M01PO")
                txtDepartment.Text = M01.Tables(0).Rows(0)("M01Department")
                txtDepartment.ReadOnly = True

                Sql = "select * from T01Delivary_Request where T01Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                M02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(M02) Then
                    txtPlanner.Text = Trim(M02.Tables(0).Rows(0)("T01Planner"))
                End If

            Else
                Search_Sales_Order = False
            End If
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""


        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
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
            .DisplayLayout.Bands(0).Columns(2).Width = 80
            .DisplayLayout.Bands(0).Columns(3).Width = 180
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
        Dim colWork As New DataColumn("##", GetType(Boolean))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = False
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

        colWork = New DataColumn("Quality", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Qty", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("New Qty", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = False

        'For I = 1 To 5
        '    colWork = New DataColumn("Del Qty" & I, GetType(Integer))
        '    '  colWork.MaxLength = 70
        '    dataTable.Columns.Add(colWork)
        '    '  colWork.ReadOnly = True
        '    colWork = New DataColumn("Del Date" & I, GetType(Date))
        '    '  colWork.MaxLength = 70
        '    dataTable.Columns.Add(colWork)

        'Next
        'colWork = New DataColumn("#", GetType(String))
        ''  colWork.MaxLength = 70
        'dataTable.Columns.Add(colWork)
        Return dataTable
    End Function

    Function Load_Gride_With_Data()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim i As Integer
        Dim T01 As DataSet
        Try
            Call Load_Gride_SalesOrder()

            i = 0

            Sql = "select T03Line_Item,sum(T03Qty_Int) as T03Qty,max(T03RefNo) as T03RefNo,max(T02Del_Req_No) as T02Del_Req_No from T03Delivary_Quat_Flutter inner join T02Delivary_Quat_Header on T02RefNo=T03RefNo  where T02OrderNo='" & Trim(cboSales_Order.Text) & "' and T02Status='A' group by T03Line_Item,T03RefNo"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            For Each DTRow4 As DataRow In M01.Tables(0).Rows
                _RefNo = M01.Tables(0).Rows(i)("T03RefNo")
                _Del_ReqNo = M01.Tables(0).Rows(i)("T02Del_Req_No")

                Dim newRow As DataRow = c_dataCustomer.NewRow
                newRow("##") = False
                newRow("Line Item") = M01.Tables(0).Rows(i)("T03Line_Item")
                Sql = "select * from M01Sales_Order_SAP where CONVERT(INT,M01Sales_Order)='" & Trim(cboSales_Order.Text) & "' and M01Line_Item='" & Trim(M01.Tables(0).Rows(i)("T03Line_Item")) & "'"
                T01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(T01) Then
                    newRow("Material") = T01.Tables(0).Rows(0)("M01Material_No")
                    newRow("Quality") = T01.Tables(0).Rows(0)("M01Quality")

                End If
                newRow("Qty") = M01.Tables(0).Rows(i)("T03Qty")
                newRow("New Qty") = M01.Tables(0).Rows(i)("T03Qty")

                c_dataCustomer.Rows.Add(newRow)


                i = i + 1
            Next

            DBEngin.CloseConnection(con)
            con.ConnectionString = ""

            '-------------------------------------------------------------


        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try


    End Function

    Private Sub cboSales_Order_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboSales_Order.AfterCloseUp
        Call Search_Sales_Order()
        Call Load_Gride_With_Data()

    End Sub

    Private Sub cboSales_Order_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboSales_Order.KeyUp
        If e.KeyCode = 13 Then
            Call Search_Sales_Order()
        End If
    End Sub

    Private Sub cboSales_Order_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboSales_Order.LostFocus
        Call Search_Sales_Order()

    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        common.ClearAll(OPR0)
        Clicked = ""
        OPR0.Enabled = True
        Call Load_Gride_SalesOrder()
    End Sub

    Private Sub chkDel_Pull_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkDel_Pull.CheckedChanged
        If chkDel_Pull.Checked = True Then
            chkDel_Push.Checked = False
            chkO_Cancel.Checked = False
            chkQ_Change.Checked = False
        End If
    End Sub

    Private Sub chkDel_Push_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkDel_Push.CheckedChanged
        If chkDel_Push.Checked = True Then
            chkDel_Pull.Checked = False
            chkO_Cancel.Checked = False
            chkQ_Change.Checked = False
        End If
    End Sub

    Private Sub chkQ_Change_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkQ_Change.CheckedChanged
        If chkQ_Change.Checked = True Then
            chkDel_Push.Checked = False
            chkO_Cancel.Checked = False
            chkDel_Pull.Checked = False
        End If

        ' UltraGrid1.Rows(0).Cells(4).Activation = Activation.AllowEdit
    End Sub

    Private Sub chkO_Cancel_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkO_Cancel.CheckedChanged
        If chkO_Cancel.Checked = True Then
            chkDel_Push.Checked = False
            chkDel_Pull.Checked = False
            chkQ_Change.Checked = False
        End If
    End Sub

    Private Sub chkAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkAll.CheckedChanged
        Dim i As Integer
        If chkAll.Checked = True Then
            i = 0
            For Each uRow As UltraGridRow In UltraGrid1.Rows
                
                UltraGrid1.Rows(i).Cells(0).Value = True

                i = i + 1
            Next
        Else
            i = 0
            For Each uRow As UltraGridRow In UltraGrid1.Rows

                UltraGrid1.Rows(i).Cells(0).Value = False

                i = i + 1
            Next
        End If
    End Sub

    Private Sub cmdUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUpdate.Click
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
        Dim Z1 As Integer
        Dim _Qty As Integer
        Dim Z2 As Integer
        Dim Z3 As Integer
        Dim Z4 As Integer
        Dim _GrideStatus As String
        Dim A As String
        Dim P01Parametere As DataSet

        Try
            If chkDel_Pull.Checked = True Then
                _GrideStatus = "DP"   '----------------------------->> Delivary Pull
            ElseIf chkDel_Push.Checked = True Then
                _GrideStatus = "DU"   '---------------------------->> Delivary Push
            ElseIf chkO_Cancel.Checked = True Then
                _GrideStatus = "OC"   '---------------------------->> Order Cancel
            End If
            '-----------------------------------------------------------------------
            'SEARCH PARAMETER
            nvcFieldList1 = "select * from P01PARAMETER where P01CODE='DR'"
            P01Parametere = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(P01Parametere) Then
                strParameter = P01Parametere.Tables(0).Rows(0)("P01NO")
            End If
            '----------------------------------------------------------------------
            

            'UPDATE PARAMETER
            nvcFieldList1 = "update P01PARAMETER set P01NO=P01NO +" & 1 & " where P01CODE='DR'"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            '----------------------------------------------------------------------
            'INSERT T06Delivary_Revision_Merchant TABLE
            i = 0
            For Each uRow As UltraGridRow In UltraGrid1.Rows
                If UltraGrid1.Rows(i).Cells(0).Text = True Then
                    nvcFieldList1 = "Insert Into T06Delivary_Revision_Merchant(T06Ref_No,T06Date,T06Criteria,T06Del_Ref,T06Sales_Order,T06Line_Item,T06Qty,T06New_Qty,T06User,T06Status)" & _
                                                                       " values(" & strParameter & ", '" & Today & "','" & _GrideStatus & "','" & _RefNo & "','" & Trim(cboSales_Order.Text) & "','" & Trim(UltraGrid1.Rows(i).Cells(1).Text) & "','" & Trim(UltraGrid1.Rows(i).Cells(4).Text) & "','" & Trim(UltraGrid1.Rows(i).Cells(5).Text) & "','" & strDisname & "','A')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                    'UPDATE T03Delivary_Quat_Flutter
                    nvcFieldList1 = "update T03Delivary_Quat_Flutter set T03Status='R' where T03RefNo='" & _RefNo & "' and T03Line_Item='" & Trim(UltraGrid1.Rows(i).Cells(1).Text) & "' and T03Status='A'"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                    'Update T01Delivary_Request
                    nvcFieldList1 = "update T01Delivary_Request set T01Status='R' where T01Sales_Order='" & Trim(cboSales_Order.Text) & "' and T01Line_Item='" & Trim(UltraGrid1.Rows(i).Cells(1).Text) & "' and T01Status='A' and T01User='" & strDisname & "'"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                End If
                i = i + 1
            Next

            transaction.Commit()
            A = MsgBox("Are you sure you want to send Revision request e-mail to Planner", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Information .........")
            If A = vbYes Then
                Call Send_Email() '------------------SENDING EMAIL

            End If
            common.ClearAll(OPR0)
            Clicked = ""
            OPR0.Enabled = True


            Call Load_Gride_SalesOrder()
            Call Load_Sales_Order()

            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""

        Catch ex As EvaluateException
            If transactionCreated = False Then transaction.Rollback()
            MessageBox.Show(Me, ex.ToString)

        Finally
            If connectionCreated Then DBEngin.CloseConnection(connection)
        End Try
    End Sub

    Function Send_Email()

        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim M02 As DataSet


        Dim OutlookApp As Microsoft.Office.Interop.Outlook._Application
        Dim wordInSubject As String
        OutlookApp = New Microsoft.Office.Interop.Outlook.Application
        Dim _RefNo1 As String
        'MsgBox(Microsoft.VisualBasic.Len(Trim(_Del_ReqNo)))
        If Microsoft.VisualBasic.Len(Trim(_Del_ReqNo)) = 1 Then
            _RefNo1 = "000" & Trim(_Del_ReqNo)
        ElseIf Microsoft.VisualBasic.Len(Trim(_Del_ReqNo)) = 2 Then
            _RefNo1 = "00" & Trim(_Del_ReqNo)
        ElseIf Microsoft.VisualBasic.Len(Trim(_Del_ReqNo)) = 3 Then
            _RefNo1 = "0" & Trim(_Del_ReqNo)
        Else
            _RefNo1 = Trim(_Del_ReqNo)
        End If

        wordInSubject = Trim(cboSales_Order.Text) & "-" & _RefNo1
        Dim scope As String = "Inbox"
        Dim filter As String = "urn:schemas:mailheader:subject LIKE '%" + wordInSubject + "%'"
        Dim advancedSearch As Microsoft.Office.Interop.Outlook.Search = Nothing
        Dim folderInbox As Microsoft.Office.Interop.Outlook.MAPIFolder = Nothing
        Dim folderSentMail As Microsoft.Office.Interop.Outlook.MAPIFolder = Nothing
        Dim ns As Microsoft.Office.Interop.Outlook.NameSpace = Nothing
        Dim oFolders As Microsoft.Office.Interop.Outlook.Folders
        Dim RootFolder As Microsoft.Office.Interop.Outlook.MAPIFolder
        Dim i As Integer

        Dim exc As New Microsoft.Office.Interop.Excel.Application
        Dim workbooks As Microsoft.Office.Interop.Excel.Workbooks = exc.Workbooks
        Dim workbook As Microsoft.Office.Interop.Excel._Workbook = workbooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet)
        Dim sheets As Microsoft.Office.Interop.Excel.Sheets = workbook.Worksheets
        Dim worksheet1 As Microsoft.Office.Interop.Excel._Worksheet = CType(sheets.Item(1), Microsoft.Office.Interop.Excel._Worksheet)
        Dim range1 As Microsoft.Office.Interop.Excel.Range


        Try
            ns = OutlookApp.GetNamespace("MAPI")
            oFolders = ns.Folders
            'MsgBox(oFolders.Count)
            i = 1
            Dim oMsg As Microsoft.Office.Interop.Outlook.MailItem
            Dim oMsg1 As Microsoft.Office.Interop.Outlook.MailItem
            Dim olFormat As OlBodyFormat

            Dim receivetime As Date
            receivetime = "1900/1/1 12:00AM"
            For i = 1 To oFolders.Count
                RootFolder = oFolders.Item(i)
                scope = "'" & RootFolder.FolderPath & "\" & "Inbox'"

                advancedSearch = OutlookApp.AdvancedSearch(scope, filter, True, advancedSearchTag)
                advancedSearch.Results.Sort("[ReceivedTime]", True)
                If advancedSearch.Results.Count > 0 Then
                    oMsg = advancedSearch.Results.GetFirst()
                    If receivetime = "1900/1/1 12:00AM" Then
                        receivetime = oMsg.ReceivedTime
                        oMsg1 = oMsg
                    Else
                        If receivetime > oMsg.ReceivedTime Then
                        Else
                            oMsg1 = oMsg
                        End If
                    End If
                End If


            Next

            '------------------------------------------------------------
            'FINDING WOORK BOOK RANGE
            'DEVELOPED BY SURANGA WIJESINGHE

            If exc.Visible = True Then
                exc.Visible = False
                exc.Visible = True
            Else
                ' exc.Visible = False
                exc.Visible = True
            End If
            worksheet1.Rows(5).Font.size = 10
            worksheet1.Rows(5).Font.bold = True
            'worksheet1.Rows(5).width = 23
            Dim A As Integer

            A = 97

            worksheet1.Rows(6).Font.size = 10
            worksheet1.Rows(6).Font.bold = True

            worksheet1.Cells(6, 1) = "Delivery Pull"
            worksheet1.Cells(6, 1).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            worksheet1.Columns("B").ColumnWidth = 14

            worksheet1.Cells(6, 2) = "Delivery Push"
            worksheet1.Cells(6, 2).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            worksheet1.Columns("C").ColumnWidth = 12

            worksheet1.Cells(6, 3) = "Quantity Change"
            worksheet1.Cells(6, 3).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            worksheet1.Columns("D").ColumnWidth = 14

            worksheet1.Cells(6, 4) = "Order Cancel"
            worksheet1.Cells(6, 4).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            worksheet1.Columns("E").ColumnWidth = 12

            worksheet1.Cells(6, 5) = "Line Item"
            worksheet1.Cells(6, 5).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            worksheet1.Columns("F").ColumnWidth = 12

            worksheet1.Cells(6, 6) = "Material"
            worksheet1.Cells(6, 6).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            worksheet1.Columns("G").ColumnWidth = 30

            worksheet1.Cells(6, 7) = "Quality"
            worksheet1.Cells(6, 7).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter

            worksheet1.Cells(6, 8) = "Qty"
            worksheet1.Cells(6, 8).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter

            worksheet1.Cells(6, 9) = "New Qty"
            worksheet1.Cells(6, 9).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter

           
            Dim x As Integer
            x = 6

            A = 97
            ' i = 0
            Dim Z As Integer
            For Z = 1 To 9
                worksheet1.Range(Chr(A) & x, Chr(A) & x).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                worksheet1.Range(Chr(A) & x, Chr(A) & x).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                worksheet1.Range(Chr(A) & x, Chr(A) & x).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                worksheet1.Range(Chr(A) & x, Chr(A) & x).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                A = A + 1
            Next

            A = 97
            For Z = 1 To 9
                worksheet1.Range(Chr(A) & x & ":" & Chr(A) & x).Interior.Color = RGB(0, 112, 192)
                A = A + 1
            Next

            x = x + 1

            Dim Y As Integer
            Y = 1

            Sql = "select max(M01Material_No) as M01Material_No,max(M01Quality) as M01Quality,T06Line_Item from T06Delivary_Revision_Merchant inner join T02Delivary_Quat_Header on T02RefNo=T06Del_Ref inner join M01Sales_Order_SAP on CONVERT(INT,M01Sales_Order)=T02OrderNo where T02OrderNo='" & Trim(cboSales_Order.Text) & "' and T06Ref_No='" & strParameter & "' group by T06Line_Item"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            For Each DTRow As DataRow In M01.Tables(0).Rows
                worksheet1.Rows(x).Font.size = 8

                worksheet1.Cells(x, Y + 4) = M01.Tables(0).Rows(i)("T06Line_Item")
                worksheet1.Cells(x, Y + 4).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                Y = Y + 1

                worksheet1.Cells(x, Y + 4) = M01.Tables(0).Rows(i)("M01Material_No")
                worksheet1.Cells(x, Y + 4).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                Y = Y + 1

                worksheet1.Cells(x, Y + 4) = M01.Tables(0).Rows(i)("M01Quality")
                worksheet1.Cells(x, Y + 4).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                Y = Y + 1

                Sql = "select * from T06Delivary_Revision_Merchant where T06Ref_No=" & strParameter & " and T06Line_Item='" & Trim(M01.Tables(0).Rows(i)("T06Line_Item")) & "' "
                M02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                Z = 0
                For Each DTRow1 As DataRow In M02.Tables(0).Rows

                    worksheet1.Cells(x, 8) = M02.Tables(0).Rows(Z)("T06Qty")
                    worksheet1.Cells(x, 8).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                    Y = Y + 1

                    worksheet1.Cells(x, 9) = M02.Tables(0).Rows(Z)("T06New_Qty")
                    worksheet1.Cells(x, 9).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                    Y = Y + 1

                    If Trim(M02.Tables(0).Rows(Z)("T06Criteria")) = "DP" Then
                        worksheet1.Cells(x, 1) = "Y"
                        worksheet1.Cells(x, 1).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                    ElseIf Trim(M02.Tables(0).Rows(Z)("T06Criteria")) = "DU" Then
                        worksheet1.Cells(x, 2) = "Y"
                        worksheet1.Cells(x, 2).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                    ElseIf Trim(M02.Tables(0).Rows(Z)("T06Criteria")) = "OC" Then
                        worksheet1.Cells(x, 4) = "Y"
                        worksheet1.Cells(x, 4).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter

                    End If
                    If Val(M02.Tables(0).Rows(Z)("T06New_Qty")) <> Val(M02.Tables(0).Rows(Z)("T06Qty")) Then
                        worksheet1.Cells(x, 3) = "Y"
                        worksheet1.Cells(x, 3).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                    End If
                    Z = Z + 1
                Next

                A = 97
                For Z = 1 To 9
                    worksheet1.Range(Chr(A) & x, Chr(A) & x).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    worksheet1.Range(Chr(A) & x, Chr(A) & x).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    worksheet1.Range(Chr(A) & x, Chr(A) & x).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    worksheet1.Range(Chr(A) & x, Chr(A) & x).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    A = A + 1
                Next

                i = i + 1
            Next

            Dim xlRn As Microsoft.Office.Interop.Excel.Range
            Dim Connect As String
            Dim strbody As String

            'strBody = "This is a test " & vbCrLf & vbCrLf & "Thanks Michael"
            '  RangetoHTML(xlRn)

            Connect = worksheet1.Range("A5:I" & x).Copy()
            xlRn = worksheet1.Range("A5:I" & x + 1)
            '_Range = "A1:" & A & lRow

            'Dim xlRn As Microsoft.Office.Interop.Excel.Range
            'Dim Connect As String
            'Dim strbody As String

            ''strBody = "This is a test " & vbCrLf & vbCrLf & "Thanks Michael"
            ''  RangetoHTML(xlRn)

            'Connect = ws.Range(_Range).Copy()
            ''SendKeys.SendWait("^V")
            'xlRn = ws.Range(_Range)
            'xlRn.Copy()
            Dim strNewText As String
            If chkDel_Pull.Checked = True Then
                strNewText = "Dear " & txtPlanner.Text & ",<br>The customer requesting delivery than you given below.<br> So please consider and re-quote the delivery "
            ElseIf chkDel_Push.Checked = True Then
                strNewText = "Dear " & txtPlanner.Text & ",<br>The customer requesting delivery are ahead you given below.<br> So please consider and re-quote the delivery "
            ElseIf chkO_Cancel.Checked = True Then
                strNewText = "Dear " & txtPlanner.Text & ",<br>The customer has cancelled the order.So please consider and advice the liability "
            End If
            Dim oResponse As MailItem
            oResponse = oMsg1.ReplyAll

            oResponse.BodyFormat = OlBodyFormat.olFormatHTML
            oResponse.HTMLBody = (strNewText & RangetoHTML(xlRn) & oResponse.HTMLBody)




            'End If
            oResponse.Display()
            ''SendKeys.SendWait("^+R")

            ''  WB.Close(False)
            ''app.Quit()


        Catch ex As EvaluateException
            MessageBox.Show(ex.Message, "An eexception is thrown")
        Finally
            If Not IsNothing(advancedSearch) Then System.Runtime.InteropServices.Marshal.ReleaseComObject(advancedSearch)
            If Not IsNothing(folderSentMail) Then System.Runtime.InteropServices.Marshal.ReleaseComObject(folderSentMail)
            If Not IsNothing(folderInbox) Then System.Runtime.InteropServices.Marshal.ReleaseComObject(folderInbox)
            If Not IsNothing(ns) Then System.Runtime.InteropServices.Marshal.ReleaseComObject(ns)
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
        TempWB.Close(SaveChanges:=False)

        'Delete the htm file we used in this function
        Kill(TempFile)

        ts = Nothing
        fso = Nothing
        TempWB = Nothing
    End Function

End Class