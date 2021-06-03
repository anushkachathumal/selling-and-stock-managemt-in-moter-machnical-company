Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Public Class frmOutstanding_Collection
    Dim _Locstatus As Boolean
    Dim _FindStatus As Boolean
    Dim c_dataCustomer1 As DataTable
    Dim c_dataCustomer2 As DataTable
    Dim _Cuscode As String
    Dim _LogStatus As Boolean
    Function Load_Customer()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet

        Try
            Sql = "select max(m06name)  as [##] from View_Outstanding_Balance inner join M06Customer_Master on M06Code=t12cus_no group by M06Code having SUM(t12cr-t12dr)>0 "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboCustomer
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 345

            End With
            con.ClearAllPools()
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
                con.close()

            End If
        End Try
    End Function

    Function Load_Grid_Collection()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select ROW_NUMBER() OVER (ORDER BY T12Cus_No ) as  ##,T12Cus_No as [Customer Code],MAX(M06Name) as [Customer Name],CAST(SUM(t12cr-t12dr) AS DECIMAL(16,2))  as [Outstanding] from View_Outstanding_Balance inner join M06Customer_Master on M06Code=T12Cus_No  group by T12Cus_No  having SUM(t12cr-t12dr)<>'0' order by T12Cus_No"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid2.DataSource = M01
            UltraGrid2.Rows.Band.Columns(0).Width = 30
            UltraGrid2.Rows.Band.Columns(1).Width = 90
            UltraGrid2.Rows.Band.Columns(2).Width = 280
            UltraGrid2.Rows.Band.Columns(3).Width = 110
            UltraGrid2.Rows.Band.Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            UltraGrid2.Rows.Band.Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Left
            con.ClearAllPools()
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
                con.close()

            End If
        End Try
    End Function

    Private Sub ExitToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem1.Click
        Me.Close()
    End Sub

    Function Load_Gride_Invo()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer2 = CustomerDataClass.MakeDataTable_Inv
        UltraGrid3.DataSource = c_dataCustomer2
        With UltraGrid3
            .DisplayLayout.Bands(0).Columns(0).Width = 80
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 90
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False

            .DisplayLayout.Bands(0).Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(9).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

        End With
    End Function

    Private Sub frmOutstanding_Collection_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        frmview_Outstanding.Close()
    End Sub

    Private Sub frmOutstanding_Collection_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_Grid_Collection()

        txtCash1.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtPayment1.ReadOnly = True
        txtPayment1.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

        txtAmount1.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

        txtPayment1.ReadOnly = True
        txtPayment1.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
        ' txtRef1.ReadOnly = True
        txtRef1.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        'txtPay.ReadOnly = True
        'txtPay.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtCus_Code1.ReadOnly = True
        txtName1.ReadOnly = True

        '  txtCash1.ReadOnly = True
        txtCash1.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        'txtAmount1.ReadOnly = True
        txtAmount1.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        '  txtChq1.ReadOnly = True

        Call Load_Gride_Invo()
        Call Load_EntryNo()
        Call Load_BANK()
        txtDate1.Text = Today

    End Sub

    Function Load_BANK()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select M11Name as [##] from M11Common where M11Status='BNK' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboBank1
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 169
                ' .Rows.Band.Columns(1).Width = 360

            End With


            con.ClearAllPools()
            con.CLOSE()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
                con.CLOSE()
            End If
        End Try
    End Function

    Function Load_EntryNo()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select * from P01Parameter where  P01Code='OU'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                If M01.Tables(0).Rows(0)("P01No") >= 1 And M01.Tables(0).Rows(0)("P01No") < 10 Then
                    txtRef1.Text = "REC-00" & M01.Tables(0).Rows(0)("P01No")
                ElseIf M01.Tables(0).Rows(0)("P01No") >= 10 And M01.Tables(0).Rows(0)("P01No") < 100 Then
                    txtRef1.Text = "REC-0" & M01.Tables(0).Rows(0)("P01No")
                Else
                    txtRef1.Text = "REC-" & M01.Tables(0).Rows(0)("P01No")
                End If
            End If

            con.ClearAllPools()
            con.CLOSE()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
                con.CLOSE()
            End If
        End Try
    End Function

    Private Sub ByCustomerToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ByCustomerToolStripMenuItem.Click
        Panel1.Visible = True
        Call Load_Customer()
        Call Load_Gride2()
        cboCustomer.Text = ""
        cboCustomer.ToggleDropdown()
    End Sub

    Private Sub RefreshToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RefreshToolStripMenuItem.Click
        Panel1.Visible = False
        _Locstatus = False
        Call Load_Grid_Collection()
    End Sub

    Function Load_Gride2()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTable_oUTSTANDING_PAY
        UltraGrid2.DataSource = c_dataCustomer1
        With UltraGrid2
            .DisplayLayout.Bands(0).Columns(0).Width = 60
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 110
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False

            .DisplayLayout.Bands(0).Columns(2).Width = 190
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(3).Width = 90
            .DisplayLayout.Bands(0).Columns(3).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(4).Width = 90
            .DisplayLayout.Bands(0).Columns(4).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(5).Width = 100
            .DisplayLayout.Bands(0).Columns(5).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(6).Width = 100
            .DisplayLayout.Bands(0).Columns(6).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(7).Width = 100
            .DisplayLayout.Bands(0).Columns(7).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(5).Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(6).Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(7).Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            ' .DisplayLayout.Bands(0).Columns(7).Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
        
            .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function

    Private Sub UltraButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton1.Click
        If Search_Customer() = True Then
            Call Load_data()
            Panel1.Visible = False
            _FindStatus = True
        Else
            MsgBox("Please select the correct customer name", MsgBoxStyle.Information, "Information .......")
        End If
    End Sub

    Function Load_data()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Dim i As Integer
        Dim _St As String
        Dim Value As Double

        Try
            Sql = "select T12inv_no,min(t12date) in_Date,max(M06Name) as M06Name,SUM(t12cr) as CR,SUM(t12dr) as DR ,SUM(t12cr-t12dr) as balance from View_Outstanding_Balance inner join M06Customer_Master on M06Code=t12cus_no where M06Name='" & Trim(cboCustomer.Text) & "' group by t12cus_no,T12inv_no  ORDER BY min(t12date) "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            Call Load_Gride2()
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                If CDbl(M01.Tables(0).Rows(i)("balance")) = 0 Then
                    i = i + 1
                    Continue For
                End If
                Dim newRow As DataRow = c_dataCustomer1.NewRow
                newRow("##") = False
                If Microsoft.VisualBasic.Left(M01.Tables(0).Rows(i)("T12inv_no"), 2) = "DS" Then
                    newRow("#Type") = "DIRECT SALES"
                ElseIf Microsoft.VisualBasic.Left(M01.Tables(0).Rows(i)("T12inv_no"), 2) = "IN" Then
                    newRow("#Type") = "JOB INVOICE"
                End If
                newRow("Customer Name") = M01.Tables(0).Rows(i)("M06Name")
                newRow("Inv.Date") = Month(M01.Tables(0).Rows(i)("in_Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("in_Date")) & "/" & Year(M01.Tables(0).Rows(i)("in_Date"))
                'newRow("Com Invoice") = M01.Tables(0).Rows(i)("T01Invoice_No")

                newRow("Inv.No") = M01.Tables(0).Rows(i)("T12inv_no")
                Value = M01.Tables(0).Rows(i)("CR")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Inv.Amount") = _St

                Value = M01.Tables(0).Rows(i)("DR")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Paid Amount") = _St

                Value = M01.Tables(0).Rows(i)("balance")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Balance") = _St


                c_dataCustomer1.Rows.Add(newRow)

                i = i + 1
            Next

            con.ClearAllPools()
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
                con.close()

            End If
        End Try
    End Function

    Private Sub ItemLookupToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ItemLookupToolStripMenuItem.Click
        _Locstatus = False
        Panel1.Visible = False
        frmview_Outstanding.Close()
        frmview_Outstanding.Show()
    End Sub

    Function Search_Customer() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select * from M06Customer_Master where M06Name='" & Trim(cboCustomer.Text) & "' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                Search_Customer = True
                _Cuscode = Trim(M01.Tables(0).Rows(0)("M06Code"))
            End If


            con.ClearAllPools()
            con.CLOSE()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
                con.CLOSE()
            End If
        End Try
    End Function

    Private Sub UltraGrid2_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles UltraGrid2.DoubleClickRow
        On Error Resume Next
        Dim i As Integer
        Dim Value As Double
        i = 0
        OPR10.Visible = True
        Call Load_Gride_Invo()
        Call Load_EntryNo()
        For Each uRow As UltraGridRow In UltraGrid2.Rows
           
            If UltraGrid2.Rows(i).Cells(0).Value = True Then
                Dim newRow As DataRow = c_dataCustomer2.NewRow
                newRow("Invo Date") = UltraGrid2.Rows(i).Cells(3).Text
                newRow("Invoice No") = UltraGrid2.Rows(i).Cells(4).Text
                newRow("Tobe Paid") = UltraGrid2.Rows(i).Cells(7).Text
                Value = Value + CDbl(UltraGrid2.Rows(i).Cells(7).Text)
                c_dataCustomer2.Rows.Add(newRow)

                txtCus_Code1.Text = _Cuscode
                txtName1.Text = UltraGrid2.Rows(i).Cells(2).Text
                ' cboStatus.Text = "M/S"
                '_Status = True
            End If
            i = i + 1
        Next
        Call Calculating_Total()
        txtCash1.Focus()
    End Sub
    Function Calculate_payment_total()
        On Error Resume Next
        Dim Value As Double

        If IsNumeric(txtCash1.Text) Then
            Value = CDbl(txtCash1.Text)
        End If

        If IsNumeric(txtAmount1.Text) Then
            Value = Value + CDbl(txtAmount1.Text)
        End If
        txtPayment1.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
        txtPayment1.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
    End Function
    Function Calculating_Total()
        On Error Resume Next
        Dim i As Integer
        Dim Value As Double

        i = 0
        For Each uRow As UltraGridRow In UltraGrid3.Rows
            Value = Value + CDbl(UltraGrid3.Rows(i).Cells(2).Text)
            i = i + 1
        Next
        txtTotal1.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
        txtTotal1.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
    End Function
    Private Sub UltraButton29_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton29.Click
        OPR10.Visible = False
    End Sub

    Private Sub txtCash1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCash1.KeyUp
        Dim Value As Double

        If e.KeyCode = 13 Then
            If txtCash1.Text <> "" Then
                If IsNumeric(txtCash1.Text) Then
                    Value = txtCash1.Text
                    txtCash1.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    txtCash1.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                End If
                cmdRecord.Focus()
            Else
                cboBank1.ToggleDropdown()
            End If
        End If
    End Sub

    Private Sub txtCash1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCash1.ValueChanged
        Call Calculate_payment_total()
    End Sub

    Private Sub cboBank1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboBank1.KeyUp
        If e.KeyCode = 13 Then
            txtChq1.Focus()
        End If
    End Sub

    Private Sub txtChq1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtChq1.KeyUp
        If e.KeyCode = 13 Then
            txtDOR.Focus()
        End If
    End Sub

    Private Sub txtDOR_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDOR.KeyUp
        If e.KeyCode = 13 Then
            txtAmount1.Focus()
        End If
    End Sub

    Private Sub txtAmount1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAmount1.KeyUp
        Dim Value As Double
        If e.KeyCode = 13 Then
            If IsNumeric(txtAmount1.Text) Then
                Value = txtAmount1.Text
                txtAmount1.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtAmount1.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            End If

            cmdRecord.Focus()
        End If
    End Sub

    Private Sub txtAmount1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAmount1.TextChanged
        Call Calculate_payment_total()
    End Sub

    Private Sub cmdRecord_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRecord.Click

        If cmdRecord.Text = "Record Payment" Then
            If txtCash1.Text <> "" Then
            Else
                txtCash1.Text = "0"
            End If

            If IsNumeric(txtCash1.Text) Then
            Else
                MsgBox("Please enter the correct cash amount", MsgBoxStyle.Information, "Information ........")
                Exit Sub

            End If

            If Trim(txtAmount1.Text) <> "" Then
            Else
                txtAmount1.Text = "0"
            End If

            If IsNumeric(txtAmount1.Text) Then
            Else
                MsgBox("Please enter the correct Chq Amount", MsgBoxStyle.Information, "Information ......")
                Exit Sub
            End If

            Call Save_Data()
        ElseIf cmdRecord.Text = "Cancel" Then


        End If
    End Sub

    Function Save_Data()
        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean
        SqlClient.SqlConnection.ClearAllPools()
        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Dim MB51 As DataSet
        Dim _GetDate As DateTime
        Dim _Get_Time As DateTime
        Dim A As String
        ' Dim B As New ReportDocument
        Dim _Paytype As String
        Dim _Balance As Double

        Dim M01 As DataSet
        Dim M02 As DataSet
        Dim I As Integer
        Dim _remark As String
        Try
            _GetDate = Month(txtDate1.Text) & "/" & Microsoft.VisualBasic.Day(txtDate1.Text) & "/" & Year(txtDate1.Text)

            _Get_Time = Month(Now) & "/" & Microsoft.VisualBasic.Day(Now) & "/" & Year(Now) & " " & Hour(Now) & ":" & Minute(Now)

            If CDbl(txtCash1.Text) > 0 And CDbl(txtAmount1.Text) > 0 Then
                _Paytype = "CASH/CHQ"
            ElseIf CDbl(txtAmount1.Text) > 0 Then
                _Paytype = "CHQUE"
            ElseIf CDbl(txtCash1.Text) > 0 Then
                _Paytype = "CASH"
            End If
            nvcFieldList1 = "select * from T15Outstanding_Collection WHERE T15Pay_No='" & Trim(txtRef1.Text) & "' AND  T15Status='PAID'"
            M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(M01) Then
                MsgBox("Record alrady exsist", MsgBoxStyle.Information, "Information .......")
                connection.ClearAllPools()
                connection.Close()
                Exit Function
            Else
                Call Load_EntryNo()
                nvcFieldList1 = "update P01Parameter set P01No=P01No+ " & 1 & " where P01Code='OU' "
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                If CDbl(txtAmount1.Text) > 0 Then
                    If Trim(cboBank1.Text) <> "" Then
                    Else
                        MsgBox("Please select the Bank Name", MsgBoxStyle.Information, "Information .......")
                        cboBank1.ToggleDropdown()
                        connection.ClearAllPools()
                        connection.Close()
                        Exit Function
                    End If
                    '==============================================================
                    If Trim(txtChq1.Text) <> "" Then
                    Else
                        MsgBox("Please enter the chq no", MsgBoxStyle.Information, "Information .......")
                        txtChq1.Focus()
                        connection.ClearAllPools()
                        connection.Close()
                        Exit Function
                    End If
                    '==============================================================
                    If Trim(txtDOR.Text) <> "" Then
                    Else
                        MsgBox("Please enter the Date of Realized", MsgBoxStyle.Information, "Information ........")
                        txtDOR.Focus()
                        connection.ClearAllPools()
                        connection.Close()
                        Exit Function
                    End If
                    '===============================================================
                    If IsDate(txtDOR.Text) Then
                    Else
                        MsgBox("Please enter the correct Date of Realized", MsgBoxStyle.Information, "Information ........")
                        txtDOR.Focus()
                        connection.ClearAllPools()
                        connection.Close()
                        Exit Function
                    End If
                    '==============================================================

                   

                    'INSERT CHQ DETAILES
                    nvcFieldList1 = "Insert Into T13Chq_Transaction(T13Ref_No,T13Tr_Type,T13Date,T13Cus_Code,T13Bank,T13Chq_No,T13DOR,T13Amount,T13Return_Status,T13Status)" & _
                                                       " values('" & Trim(txtRef1.Text) & "','OUTSTANDING_COLLECTION', '" & _GetDate & "','" & Trim(txtCus_Code1.Text) & "','" & Trim(cboBank1.Text) & "','" & Trim(txtChq1.Text) & "','" & (txtDOR.Text) & "','" & txtAmount1.Text & "','-','A')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)


                End If
                '===================================================================
                'outstanding summery
                nvcFieldList1 = "Insert Into T16Outstanding_Pay_Summery(T16Pay_No,T16Date,T16Net_Amount,T16Cash,T16Chq,T16Status)" & _
                                                    " values('" & Trim(txtRef1.Text) & "', '" & _GetDate & "','" & CDbl(txtPayment1.Text) & "','" & CDbl(txtCash1.Text) & "','" & CDbl(txtAmount1.Text) & "','A')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                '====================================================================
                'PAYMENT SUMMERY
                nvcFieldList1 = "Insert Into T11Income_Summery(T11Invo_No,T11Tr_Type,T11Job_No,T11Date,T11Net_Amount,T11Cash,T11Credit,T11Chq,T11Card,T11Status)" & _
                                                    " values('" & Trim(txtRef1.Text) & "', 'OUTSTANDING_COLLECTION','-','" & _GetDate & "','" & CDbl(txtPayment1.Text) & "','" & CDbl(txtCash1.Text) & "','0','" & CDbl(txtAmount1.Text) & "','0','A')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                '====================================================================
                'PROFT & LOSS
                _remark = "Outstandng Collection - " & Trim(txtRef1.Text)
                nvcFieldList1 = "Insert Into T04Profit_Loss(T04Ref_No,T04Tr_Type,T04Date,T04Description,T04Cr,T04Dr,T04Status)" & _
                                                 " values('" & Trim(txtRef1.Text) & "', 'OUTSTANDING_COLLECTION','" & _GetDate & "','" & _remark & "','" & CDbl(txtPayment1.Text) & "','0','A')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                _Balance = txtPayment1.Text

                I = 0
                For Each uRow As UltraGridRow In UltraGrid3.Rows
                    If CDbl(UltraGrid3.Rows(I).Cells(2).Value) >= _Balance Then
                        nvcFieldList1 = "Insert Into T15Outstanding_Collection(T15Pay_No,T15Inv_No,T15Date,T15Cus_Code,T15Pay_Type,T15Amount,T15Status,T15Tobe_Paid)" & _
                                             " values('" & Trim(txtRef1.Text) & "', '" & Trim(UltraGrid3.Rows(I).Cells(1).Text) & "','" & _GetDate & "','" & Trim(txtCus_Code1.Text) & "','" & _Paytype & "','" & _Balance & "','PAID','" & Trim(UltraGrid3.Rows(I).Cells(2).Text) & "')"
                        ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                        _Balance = _Balance - CDbl(UltraGrid3.Rows(I).Cells(2).Text)
                        Exit For
                    Else
                        nvcFieldList1 = "Insert Into T15Outstanding_Collection(T15Pay_No,T15Inv_No,T15Date,T15Cus_Code,T15Pay_Type,T15Amount,T15Status,T15Tobe_Paid)" & _
                                                                " values('" & Trim(txtRef1.Text) & "', '" & Trim(UltraGrid3.Rows(I).Cells(1).Text) & "','" & _GetDate & "','" & Trim(txtCus_Code1.Text) & "','" & _Paytype & "','" & CDbl(UltraGrid3.Rows(I).Cells(2).Text) & "','PAID','" & Trim(UltraGrid3.Rows(I).Cells(2).Text) & "')"
                        ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                        _Balance = _Balance - CDbl(UltraGrid3.Rows(I).Cells(2).Text)
                    End If
                    I = I + 1
                Next

                If _Balance > 0 Then
                    nvcFieldList1 = "Insert Into T15Outstanding_Collection(T15Pay_No,T15Inv_No,T15Date,T15Cus_Code,T15Pay_Type,T15Amount,T15Status)" & _
                                                               " values('" & Trim(txtRef1.Text) & "', '-','" & _GetDate & "','" & Trim(txtCus_Code1.Text) & "','" & _Paytype & "','" & _Balance & "','PAID')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                End If

                MsgBox("Record update Successfully", MsgBoxStyle.Information, "Information .......")
                transaction.Commit()
            End If

            connection.ClearAllPools()
            connection.Close()
            Call Clear_text()
            Call Load_Grid_Collection()
            OPR10.Visible = False
        Catch returnMessage As Exception
            transaction.Rollback()
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.ClearAllPools()
                connection.Close()
            End If
        End Try
    End Function
    Function Search_Records()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Dim i As Integer
        Dim Value As Double
        Dim _St As String
        Dim _Total As Double
        Dim _DOR As String

        Try

            Sql = "select * from T15Outstanding_Collection inner join M06Customer_Master on T15Cus_Code=M06Code  where T15Status='A' and T15Pay_No='" & Trim(txtRef1.Text) & "' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                txtCus_Code1.Text = Trim(M01.Tables(0).Rows(0)("T15Cus_Code"))
                txtName1.Text = Trim(M01.Tables(0).Rows(0)("M06Name"))
                txtDate1.Text = Trim(M01.Tables(0).Rows(0)("T15Date"))
            End If

            Call Load_Gride_Invo()
            Sql = "select * from T15Outstanding_Collection inner join T12OutStanding on T12Inv_No=T15Inv_No where T15Pay_No='" & Trim(txtRef1.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            For Each DTRow2 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer2.NewRow
                newRow("Invo Date") = Month(M01.Tables(0).Rows(i)("T12Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T12Date")) & "/" & Year(M01.Tables(0).Rows(i)("T12Date"))
                newRow("Invoice No") = Trim(M01.Tables(0).Rows(i)("T15Inv_No"))
                Value = Trim(M01.Tables(0).Rows(i)("T15Tobe_Paid"))
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Tobe Paid") = _St
                _Total = _Total + CDbl(M01.Tables(0).Rows(i)("T15Tobe_Paid"))
                c_dataCustomer2.Rows.Add(newRow)
                i = i + 1
            Next
            txtTotal1.Text = (_Total.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            txtTotal1.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", _Total))

            Sql = "select * from T16Outstanding_Pay_Summery where T16Pay_No='" & Trim(txtRef1.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                Value = M01.Tables(0).Rows(0)("T16Cash")
                txtCash1.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtCash1.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            End If
            '==============================================================================================================
            Sql = "select * from T13Chq_Transaction where T13Ref_No='" & Trim(txtRef1.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                cboBank1.Text = M01.Tables(0).Rows(0)("T13Bank")
                txtChq1.Text = M01.Tables(0).Rows(0)("T13Chq_No")
                _DOR = Year(M01.Tables(0).Rows(0)("T13DOR"))
                If Month(M01.Tables(0).Rows(0)("T13DOR")) > 10 Then
                    _DOR = _DOR & Month(M01.Tables(0).Rows(0)("T13DOR"))
                Else
                    _DOR = _DOR & "0" & Month(M01.Tables(0).Rows(0)("T13DOR"))
                End If

                If Microsoft.VisualBasic.Day(M01.Tables(0).Rows(0)("T13DOR")) > 10 Then
                    _DOR = _DOR & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(0)("T13DOR"))
                Else
                    _DOR = _DOR & "0" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(0)("T13DOR"))
                End If
                txtDOR.Text = _DOR
                Value = M01.Tables(0).Rows(0)("T13Amount")
                txtAmount1.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtAmount1.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            End If

            con.ClearAllPools()
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
                con.close()

            End If
        End Try
    End Function
    Function Clear_text()
        Me.txtCash1.Text = ""
        Me.txtChq1.Text = ""
        Me.cboBank1.Text = ""
        Me.txtPayment1.Text = ""
        Me.txtDOR.Text = ""
        Me.txtCus_Code1.Text = ""
        Me.txtName1.Text = ""
        Me.txtNaration.Text = ""
        Me.txtAmount1.Text = ""
        cmdRecord.Text = "Record Payment"
        OPR10.Visible = False
        _Locstatus = False
        Call Load_Gride2()
        Call Load_Gride2()

    End Function

    Private Sub UltraButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton3.Click
        Call Clear_text()
        OPR10.Visible = False
    End Sub
End Class