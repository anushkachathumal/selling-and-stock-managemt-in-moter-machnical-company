'// Sales ordering module for the Planner
'// Development Date - 09.15.2014
'// Developed by - Suranga Wijesinghe
'// Audit by     - Amila Priyankara - TJL
'// Referance Table - - T01Delivary_Request

'//---------------------------------------------------------->>>
'Automate the Email  send by merchant to  Planner & Excell data migration
'once merchant enter the S order 2 system.system will gather all requred infor by callling SAP sales order file & fill the e mail requrment

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



Public Class frmDel_Cus
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim c_dataCustomer1 As DataTable
    Dim _RefNo As String
    Dim _T01RefNo As Integer

    Function Load_Gride_SalesOrder()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = MakeDataTable_Delivary_Quatation()
        UltraGrid1.DataSource = c_dataCustomer1
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 50
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 80
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 210
            '.DisplayLayout.Bands(0).Columns(3).Width = 60
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

        colWork = New DataColumn("SH/Line No", GetType(Integer))
        'colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        ' colWork.ReadOnly = True

        colWork = New DataColumn("Pln Qty", GetType(Integer))
        '  colWork.MaxLength = 70
        dataTable.Columns.Add(colWork)
        '  colWork.ReadOnly = True
        colWork = New DataColumn("Pln Date", GetType(String))
        '  colWork.MaxLength = 70
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("D Date to Customer", GetType(Date))
        '  colWork.MaxLength = 70
        dataTable.Columns.Add(colWork)


        
        'colWork = New DataColumn("#", GetType(String))
        ''  colWork.MaxLength = 70
        'dataTable.Columns.Add(colWork)
        Return dataTable
    End Function


    Private Sub frmDel_Cus_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_Gride_SalesOrder()
        Call Load_Sales_Order()

    End Sub

    Function Load_Sales_Order()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        'Load sales order to cboSO combobox

        Try
            Sql = "select T02OrderNo as [Sales Order] from T02Delivary_Quat_Header " & _
                  " inner join T03Delivary_Quat_Flutter on T02RefNo=T03RefNo inner join T01Delivary_Request on " & _
                  " T01Sales_Order=T02OrderNo where T01User='" & strDisname & "' and T03FD_Status='N' group by T02OrderNo "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboSO
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 180
                '  .Rows.Band.Columns(1).Width = 260
                '

            End With
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            'With txtNSL
            '    .DataSource = M01
            '    .Rows.Band.Columns(0).Width = 225
            'End With

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Function Search_Salrs_Order() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim i As Integer

        Try
            'SERCH SALES ORDER
            Sql = "select T01RefNo,T01Sales_Order ,max(M01Cuatomer_Name) as [Customer] from T01Delivary_Request inner join M01Sales_Order_SAP on T01Sales_Order=CONVERT(INT,M01Sales_Order) where T01User='" & strDisname & "'  and T01Sales_Order='" & Trim(cboSO.Text) & "' group by T01Sales_Order,T01RefNo"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then

                _RefNo = M01.Tables(0).Rows(0)("T01RefNo")
                txtCustomer.Text = M01.Tables(0).Rows(0)("Customer")
                Search_Salrs_Order = True


            Else
                Search_Salrs_Order = False
            End If
            '----------------------------------------------------------------------------------

            Call Load_Gride_SalesOrder()
            Call Load_Data_Gride()
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub cboSO_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboSO.AfterCloseUp
        Call Search_Salrs_Order()
    End Sub

    Private Sub cboSO_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboSO.TextChanged
        Call Search_Salrs_Order()
    End Sub

    Function Load_Data_Gride()
        Dim Sql As String
        Dim con = New SqlConnection()
        Dim con1 = New SqlConnection()

        con = DBEngin.GetConnection(True)
        con1 = DBEngin1.GetConnection(True)
        Dim M01 As DataSet
        Dim T02 As DataSet

        Dim i As Integer
        Dim Value As Double
        Dim _Qty As Double

        Try
            'Search Data to M01Sales_Order_SAP table
            Call Load_Gride_SalesOrder()
            Sql = "select T02OrderNo,M01Line_Item,M01Material_No,M01Quality,M01SO_Qty from T02Delivary_Quat_Header " & _
                  " inner join M01Sales_Order_SAP on T02OrderNo=CONVERT(INT,M01Sales_Order) where T02Status='A' and T02OrderNo='" & Trim(cboSO.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            Dim y As Integer

            For Each DTRow1 As DataRow In M01.Tables(0).Rows

                Dim newRow As DataRow = c_dataCustomer1.NewRow
                _Qty = 0
                newRow("Line Item") = M01.Tables(0).Rows(i)("M01Line_Item")
                newRow("Material") = M01.Tables(0).Rows(i)("M01Material_No")
                newRow("Quality") = M01.Tables(0).Rows(i)("M01Quality")
                Value = M01.Tables(0).Rows(i)("M01SO_Qty")
                _Qty = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _Qty = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Qty") = _Qty
                y = 0
                Sql = "select T03Qty_Int,T03Date,T02RefNo from T03Delivary_Quat_Flutter inner join T02Delivary_Quat_Header on T03RefNo=T02RefNo where T02Del_Req_No=" & _RefNo & " and T03Status='A' and T03FD_Status='N' and T02Status='A'"
                T02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                For Each DTRow2 As DataRow In T02.Tables(0).Rows
                    _T01RefNo = T02.Tables(0).Rows(y)("T02RefNo")
                    If i = 0 And y = 0 Then
                        newRow("SH/Line No") = y + 1
                        newRow("Pln Qty") = T02.Tables(0).Rows(y)("T03Qty_Int")
                        newRow("Pln Date") = T02.Tables(0).Rows(y)("T03Date")
                        c_dataCustomer1.Rows.Add(newRow)
                    Else
                        Dim newRow1 As DataRow = c_dataCustomer1.NewRow
                        newRow1("SH/Line No") = y + 1
                        newRow1("Pln Qty") = T02.Tables(0).Rows(y)("T03Qty_Int")
                        newRow1("Pln Date") = T02.Tables(0).Rows(y)("T03Date")
                        c_dataCustomer1.Rows.Add(newRow1)

                    End If
                    y = y + 1
                Next


                'newRow("PP") = False


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

    Private Sub cboSO_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles cboSO.InitializeLayout

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
        Dim M01 As DataSet
        Dim I As Integer
        Dim A As String
        Dim _LineItem As String

        Try
            '        If Search_Salrs_Order() = True Then
            '        Else
            '            Dim result1 As String
            '            result1 = MessageBox.Show("Please select the Sales Order ", "Information .....", _
            'MessageBoxButtons.OK, MessageBoxIcon.Information)
            '            If result1 = Windows.Forms.DialogResult.OK Then
            '                cboSO.ToggleDropdown()
            '                Exit Sub
            '            End If
            '        End If
            nvcFieldList1 = "select * from T02Delivary_Quat_Header where T02OrderNo='" & Trim(cboSO.Text) & "' and T02Status='A'"
            dsUser = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(dsUser) Then
            Else
                Dim result1 As String
                result1 = MessageBox.Show("Please select the Sales Order ", "Information .....", _
    MessageBoxButtons.OK, MessageBoxIcon.Information)
                If result1 = Windows.Forms.DialogResult.OK Then
                    cboSO.ToggleDropdown()
                    Exit Sub
                End If
            End If
            '----------------------------------------------------------------------------------------
            I = 0
            For I = 0 To UltraGrid1.Rows.Count - 1
                If Trim((UltraGrid1.Rows(I).Cells(0).Text)) <> "" Then
                    _LineItem = Trim((UltraGrid1.Rows(I).Cells(0).Text))
                End If

                If IsDate(Trim((UltraGrid1.Rows(I).Cells(7).Text))) Then
                    nvcFieldList1 = "update T03Delivary_Quat_Flutter set T03Final_Delivary='" & Trim((UltraGrid1.Rows(I).Cells(7).Text)) & "',T03FD_Status='Y' where T03RefNo=" & _T01RefNo & " and T03Line_Item='" & _LineItem & "' and T03Qty_Int=" & Trim((UltraGrid1.Rows(I).Cells(5).Text)) & " and T03Date='" & Trim((UltraGrid1.Rows(I).Cells(6).Text)) & "'"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                End If
                ' I = I + 1
            Next
            transaction.Commit()
            Dim result2 As String
            result2 = MessageBox.Show("Record successfully updated ", "Information .....", _
MessageBoxButtons.OK, MessageBoxIcon.Information)
            If result2 = Windows.Forms.DialogResult.OK Then
                common.ClearAll(OPR0)
                Clicked = ""
                OPR0.Enabled = True
                Call Load_Gride_SalesOrder()
                Call Load_Sales_Order()
            End If
          

            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""

        Catch ex As EvaluateException
            If transactionCreated = False Then transaction.Rollback()
            MessageBox.Show(Me, ex.ToString)

        Finally
            If connectionCreated Then DBEngin.CloseConnection(connection)
        End Try

    End Sub

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub UltraButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton1.Click
        Me.Close()
    End Sub
End Class