
Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors

Public Class frmYarn_Request_Conformation
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim c_dataCustomer1 As System.Data.DataTable
    Dim c_dataCustomer2 As System.Data.DataTable

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Private Sub frmYarn_Request_Conformation_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtPlanner.ReadOnly = True
        Call Load_Sales_Order()
        Call Load_Gride()
        Call Load_Gride1()
        Call Load_Data_Grige1()
        txtPlanner.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
    End Sub

    Function MakeDataTable_Yarn_Request() As DataTable
        Dim I As Integer
        Dim X As Integer
        Dim _Lastweek As Integer


        ' MsgBox(DatePart("ww", Today))
        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Ref.No", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Yarn", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Required Date", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Qty", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Receiving Date", GetType(Date))
        ' colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = False

       

        Return dataTable
    End Function

    Function MakeDataTable_Request() As DataTable
        Dim I As Integer
        Dim X As Integer
        Dim _Lastweek As Integer


        ' MsgBox(DatePart("ww", Today))
        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Planner’s Name", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Sales Order", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Line Item", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Req Date", GetType(String))
        ' colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Days", GetType(String))
        ' colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Hour", GetType(String))
        ' colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Minute", GetType(String))
        ' colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True


        Return dataTable
    End Function

    Function Load_Gride()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = MakeDataTable_Yarn_Request()
        UltraGrid2.DataSource = c_dataCustomer1
        With UltraGrid2
            .DisplayLayout.Bands(0).Columns(0).Width = 80
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 210
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 90
            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(3).Width = 100
            .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(3).Width = 60
            '.DisplayLayout.Bands(0).Columns(5).Width = 60
            '.DisplayLayout.Bands(0).Columns(8).Width = 60
            '.DisplayLayout.Bands(0).Columns(7).Width = 70
            '.DisplayLayout.Bands(0).Columns(9).Width = 60

        End With
    End Function

    Function Load_Gride1()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer2 = MakeDataTable_Request()
        UltraGrid1.DataSource = c_dataCustomer2
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 70
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 60
            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(3).Width = 60
            .DisplayLayout.Bands(0).Columns(4).Width = 50
            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(5).Width = 50
            .DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(6).Width = 60
            .DisplayLayout.Bands(0).Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            '.DisplayLayout.Bands(0).Columns(3).Width = 60
            '.DisplayLayout.Bands(0).Columns(5).Width = 60
            '.DisplayLayout.Bands(0).Columns(8).Width = 60
            '.DisplayLayout.Bands(0).Columns(7).Width = 70
            '.DisplayLayout.Bands(0).Columns(9).Width = 60

        End With
    End Function

    Function Load_Sales_Order()
        Dim i As Integer
        Dim vcWhere As String
        Dim M01 As DataSet
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)

        Try
            vcWhere = "T14Status='N'  "
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "YRS"), New SqlParameter("@vcWhereClause1", vcWhere))
            With cboSO
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 160
                '.Rows.Band.Columns(1).Width = 260
            End With
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

    Function Load_LineItem()
        Dim i As Integer
        Dim vcWhere As String
        Dim M01 As DataSet
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)

        Try
            vcWhere = "T14Status='N' and T14Sales_order='" & Trim(cboSO.Text) & "' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "YRL"), New SqlParameter("@vcWhereClause1", vcWhere))
            With cboLine_Item
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 80
                '.Rows.Band.Columns(1).Width = 260
            End With
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

    Function Load_Data_Grige1()
        Dim i As Integer
        Dim vcWhere As String
        Dim M01 As DataSet
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim Diff As TimeSpan
        Dim Value As Double
        Dim _STValue As String

        Try
            vcWhere = "T14Status='N' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "YRQ"), New SqlParameter("@vcWhereClause1", vcWhere))
            i = 0
            For Each DTRow4 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer2.NewRow

                newRow("Planner’s Name") = M01.Tables(0).Rows(i)("T14Req_By")
                newRow("Sales Order") = M01.Tables(0).Rows(i)("T14Sales_order")
                newRow("Line Item") = M01.Tables(0).Rows(i)("T14Line_Item")

                If IsDate(M01.Tables(0).Rows(i)("T14Req_Date")) Then
                    newRow("Req Date") = Month(M01.Tables(0).Rows(i)("T14Req_Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T14Req_Date")) & "/" & Year(M01.Tables(0).Rows(i)("T14Req_Date"))
                End If
                ' _To = Month(M02.Tables(0).Rows(Z)("M33Date")) & "/" & Microsoft.VisualBasic.Day(M02.Tables(0).Rows(Z)("M33Date")) & "/" & Year(M02.Tables(0).Rows(Z)("M33Date"))
                Diff = Now.Subtract(M01.Tables(0).Rows(i)("T14Time"))
                newRow("Days") = Diff.Days
                newRow("Hour") = Diff.Hours
                newRow("Minute") = Diff.Minutes

                'If Diff.Days < 30 Then
                '    newRow("Age") = "Below 1 Month"
                'ElseIf Diff.Days >= 30 And Diff.Days < 60 Then
                '    newRow("Age") = "Below 2 Month"
                'ElseIf Diff.Days >= 60 And Diff.Days < 90 Then
                '    newRow("Age") = "Below 3 Month"
                'Else
                '    newRow("Age") = "above 3 Month"
                'End If

                c_dataCustomer2.Rows.Add(newRow)

                i = i + 1
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

    Private Sub cboSO_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboSO.AfterCloseUp
        Call Load_LineItem()
    End Sub

    Private Sub cboSO_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboSO.LostFocus
        Call Load_LineItem()
    End Sub

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        cboLine_Item.Text = ""
        cboSO.Text = ""
        txtPlanner.Text = ""
        Call Load_Gride()
        Call Load_Sales_Order()
        cboSO.ToggleDropdown()

    End Sub

    Function Search_Records()
        Dim i As Integer
        Dim vcWhere As String
        Dim M01 As DataSet
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim Value As Double
        Dim _STValue As String

        Try
            vcWhere = "T14Status='N' and T14Sales_order='" & Trim(cboSO.Text) & "' and T14Line_Item=" & cboLine_Item.Text & " "
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "YRQ"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                txtPlanner.Text = M01.Tables(0).Rows(0)("T14Req_By")
                i = 0
                For Each DTRow4 As DataRow In M01.Tables(0).Rows
                    Dim newRow As DataRow = c_dataCustomer1.NewRow
                    newRow("Ref.No") = M01.Tables(0).Rows(i)("T14Ref_no")
                    newRow("Yarn") = M01.Tables(0).Rows(i)("T14Yarn")
                    If IsDate(M01.Tables(0).Rows(i)("T14Req_Date")) Then
                        newRow("Required Date") = Month(M01.Tables(0).Rows(i)("T14Req_Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T14Req_Date")) & "/" & Year(M01.Tables(0).Rows(i)("T14Req_Date"))
                    Else
                        newRow("Required Date") = "Week - " & M01.Tables(0).Rows(i)("T14Week")
                    End If
                    Value = M01.Tables(0).Rows(i)("T14Available")
                    _STValue = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    _STValue = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    newRow("Qty") = _STValue
                    c_dataCustomer1.Rows.Add(newRow)

                    i = i + 1
                Next

            End If
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

    Private Sub cboLine_Item_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboLine_Item.AfterCloseUp
        Call Search_Records()
    End Sub

    Private Sub cboLine_Item_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles cboLine_Item.InitializeLayout

    End Sub

    Private Sub cboLine_Item_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboLine_Item.LostFocus
        ' Call Search_Records()
    End Sub

    Private Sub UltraButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton6.Click
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
            vcWhere = "T14Status='N' and T14Sales_order='" & Trim(cboSO.Text) & "' and T14Line_Item=" & cboLine_Item.Text & " "
            M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "YRQ"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                i = 0
                For Each uRow As UltraGridRow In UltraGrid2.Rows
                    If IsDate(UltraGrid2.Rows(i).Cells(4).Value) Then
                        nvcFieldList1 = "UPDATE T14Yarn_Request SET T14Status='Y',T14Received_Date='" & UltraGrid2.Rows(i).Cells(4).Value & "',T14Confirm_Date='" & Now & "',T14Procument_User='" & strDisname & "' WHERE T14Sales_order='" & cboSO.Text & "' AND T14Line_Item='" & cboLine_Item.Text & "' and T14Ref_no='" & Trim(UltraGrid2.Rows(i).Cells(0).Value) & "' "
                        ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                    End If
                    i = i + 1
                Next

            End If

            MsgBox("Records update successfully", MsgBoxStyle.Information, "Information ......")
            transaction.Commit()
            Call Load_Sales_Order()
            Call Load_Gride()
            Call Load_Gride1()
            Call Load_Data_Grige1()
            cboSO.Text = ""
            cboLine_Item.Text = ""
            cboSO.ToggleDropdown()

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)

                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
                connection.Close()
            End If
        End Try
    End Sub

    Private Sub cboSO_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles cboSO.InitializeLayout

    End Sub
End Class