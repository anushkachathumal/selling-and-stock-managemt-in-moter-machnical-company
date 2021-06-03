Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports CrystalDecisions.CrystalReports.Engine
Public Class frmPO
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim c_dataCustomer1 As DataTable
    Dim _Category As String
    Dim _Comcode As String
    Dim _Loccode As String
    Dim _FromLocCode As String
    Dim _EntryNo As Integer

    Private Sub frmPO_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        frmViewPO.Close()
    End Sub

    Private Sub frmPO_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyUp
        If e.KeyCode = Keys.F2 Then
            frmViewPO.Show()
        End If
    End Sub

    Private Sub frmPO_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        _Comcode = ConfigurationManager.AppSettings("LOCCode")
        txtDate.Text = Today
        txtQty.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtTotal.ReadOnly = True
        txtRate.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtTotal.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
        txtNett.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
        txtNett.ReadOnly = True

        Call Load_Combo()
        Call Load_EntryNo()
        Call Load_Item_Name()
        Call Load_Item_Code()
        Call Load_Gride2()
        txtEntry.ReadOnly = True
        Call Load_Gride_Item()
    End Sub

    Function Load_EntryNo()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select * from P01Parameter where  P01Code='PO'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                If M01.Tables(0).Rows(0)("P01LastNo") >= 1 And M01.Tables(0).Rows(0)("P01LastNo") < 10 Then
                    txtEntry.Text = "PO-00" & M01.Tables(0).Rows(0)("P01LastNo")
                ElseIf M01.Tables(0).Rows(0)("P01LastNo") >= 10 And M01.Tables(0).Rows(0)("P01LastNo") < 100 Then
                    txtEntry.Text = "PO-0" & M01.Tables(0).Rows(0)("P01LastNo")
                Else
                    txtEntry.Text = "PO-" & M01.Tables(0).Rows(0)("P01LastNo")
                End If
            End If


            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Function Load_Combo()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M09Name as [##] from M09Supplier where M09Active='A' and M09Loc_Code='" & _Comcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboLocation
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 270
                '  .Rows.Band.Columns(1).Width = 160


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
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
            End If
        End Try
    End Function

    Function Search_Records()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Dim Value As Double
        Dim _St As String
        Dim i As Integer
        Try
            Sql = "select * from View_PO where T12PO_no='" & txtEntry.Text & "' and T12Loc_Code='" & _Comcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                txtRemark.Text = Trim(M01.Tables(0).Rows(0)("Remark"))
                txtDate.Text = Trim(M01.Tables(0).Rows(0)("date"))
                cboLocation.Text = Trim(M01.Tables(0).Rows(0)("Supplier"))
                Value = Trim(M01.Tables(0).Rows(0)("Total"))
                txtNett.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtNett.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                _EntryNo = M01.Tables(0).Rows(0)("T12Ref_No")
            End If

            Call Load_Gride2()
            i = 0
            Sql = "select * from View_PO_Fluter where T13Ref_No=" & _EntryNo & ""
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow
                newRow("Item Code") = Trim(M01.Tables(0).Rows(i)("T13Item_Code"))
                newRow("Item Name") = Trim(M01.Tables(0).Rows(i)("M03Item_Name"))
                Value = Trim(M01.Tables(0).Rows(i)("T13Rate"))
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Rate") = _St
                newRow("Qty") = Trim(M01.Tables(0).Rows(i)("T13Qty"))
                Value = CDbl(M01.Tables(0).Rows(i)("T13Rate")) * CDbl(M01.Tables(0).Rows(i)("T13Qty"))
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Total") = _St
                c_dataCustomer1.Rows.Add(newRow)
                i = i + 1
            Next
            con.close()



        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function
    Function Load_Gride_Item3()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim dsUser As DataSet
        Try
            Sql = "select M03Item_Code as [Item Code],M03Item_Name as [Item Name],CONVERT(varchar,CAST(M03Retail_Price AS money), 1) as [Retail Price] from M03Item_Master where M03Item_Name  like '%" & txtFind.Text & "%' and M03Status='A' and M03Com_Code='" & _Comcode & "' order by M03Item_Code"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid3.DataSource = dsUser
            UltraGrid3.Rows.Band.Columns(0).Width = 130
            UltraGrid3.Rows.Band.Columns(1).Width = 370
            UltraGrid3.Rows.Band.Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function
    Function Load_Item_Code()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M03Item_Code as [Item Code] from M03Item_Master where M03Status='A' and M03Com_Code='" & _Comcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboCode
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 130
                '  .Rows.Band.Columns(1).Width = 160


            End With
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Function Load_Item_Name()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M03Item_Name as [Item Name] from M03Item_Master where M03Status='A' and M03Com_Code='" & _Comcode & "' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboItemName
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 370
                '  .Rows.Band.Columns(1).Width = 160


            End With
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub


    Function Load_Gride2()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTable_PO
        UltraGrid1.DataSource = c_dataCustomer1
        With UltraGrid1
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

    Function Search_ItemName() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim Value As Double
        Try
            Sql = "select * from M03Item_Master where  M03Item_Code='" & Trim(cboCode.Text) & "' and m03status='A' and M03Com_Code='" & _Comcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                Search_ItemName = True
                cboItemName.Text = M01.Tables(0).Rows(0)("M03Item_Name")
                Value = M01.Tables(0).Rows(0)("M03Cost_Price")
                txtRate.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtRate.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            End If
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Function Clear_Text()
        Me.cboItemName.Text = ""
        Me.cboItemName.Text = ""
        Me.txtRemark.Text = ""
        Me.cboLocation.Text = ""
        Me.txtRate.Text = ""
        Me.txtTotal.Text = ""
        Me.txtNett.Text = ""
        Call Load_Gride2()
        Call Load_EntryNo()
    End Function

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        Call Clear_Text()
        cboLocation.ToggleDropdown()
        Call Load_EntryNo()
    End Sub

    Function Load_Gride_Item()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim dsUser As DataSet
        Try
            Sql = "select M03Item_Code as [Item Code],M03Item_Name as [Item Name] from M03Item_Master where m03Status='A' and M03Com_Code='" & _Comcode & "' order by M03Item_Code"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid3.DataSource = dsUser
            UltraGrid3.Rows.Band.Columns(0).Width = 130
            UltraGrid3.Rows.Band.Columns(1).Width = 370
            ' UltraGrid1.Rows.Band.Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub cboCode_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboCode.KeyUp
        If e.KeyCode = 13 Then
            Call Search_ItemName()
            txtRate.Focus()
        ElseIf e.KeyCode = Keys.F1 Then
            OPR5.Visible = True
            txtFind.Focus()
        ElseIf e.KeyCode = Keys.F2 Then
            frmViewPO.Show()
        ElseIf e.KeyCode = Keys.Escape Then
            OPR5.Visible = False
            frmViewPO.Close()
            txtFind.Text = ""
        End If
    End Sub

    Function Search_ItemCode()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim Value As Double

        Try
            Sql = "select * from M03Item_Master where  M03Item_Name='" & Trim(cboItemName.Text) & "' and m03Status='A' and M03Com_Code='" & _Comcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                cboCode.Text = M01.Tables(0).Rows(0)("M03Item_Code")
                Value = M01.Tables(0).Rows(0)("M03Cost_Price")
                txtRate.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtRate.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            End If
            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
               con.close()
            End If
        End Try
    End Function

    Private Sub cboItemName_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboItemName.KeyUp
        If e.KeyCode = 13 Then
            Call Search_ItemCode()
            txtRate.Focus()
        ElseIf e.KeyCode = Keys.F1 Then
            OPR5.Visible = True
            txtFind.Focus()
        ElseIf e.KeyCode = Keys.F2 Then
            frmViewPO.Show()
        ElseIf e.KeyCode = Keys.Escape Then
            OPR5.Visible = False
            txtFind.Text = ""
            frmViewPO.Close()
        End If
    End Sub


    Private Sub txtQty_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtQty.KeyUp
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim Value As Double

        Try
            If e.KeyCode = 13 Then
                If txtQty.Text <> "" Then
                    If IsNumeric(txtQty.Text) Then
                    Else
                        MsgBox("Please enter the correct Qty", MsgBoxStyle.Information, "Information .........")
                        txtQty.Focus()
                        Exit Sub
                    End If
                    If txtRate.Text <> "" Then
                    Else
                        txtRate.Text = "0"
                    End If

                    Call Calculation()

                    Sql = "select * from M03Item_Master where M03Item_Code='" & cboCode.Text & "' and M03Status='A' and M03Com_Code='" & _Comcode & "'"
                    M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(M01) Then
                    Else
                        MsgBox("Please select the correct Item code", MsgBoxStyle.Information, "Information ......")
                        cboCode.ToggleDropdown()
                        Exit Sub
                    End If
                   
                    Dim newRow As DataRow = c_dataCustomer1.NewRow
                    newRow("Item Code") = Trim(cboCode.Text)
                    newRow("Item Name") = cboItemName.Text
                    newRow("Rate") = txtRate.Text
                    newRow("Qty") = txtQty.Text
                    newRow("Total") = txtTotal.Text
                    c_dataCustomer1.Rows.Add(newRow)

                    If txtNett.Text <> "" Then
                    Else
                        txtNett.Text = "0"
                    End If
                    Value = CDbl(txtNett.Text) + CDbl(txtTotal.Text)
                    txtNett.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    txtNett.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                    txtQty.Text = ""
                    cboCode.Text = ""
                    cboItemName.Text = ""
                    txtTotal.Text = ""
                    txtRate.Text = ""
                    cboCode.Focus()
                End If
            ElseIf e.KeyCode = Keys.F2 Then
                frmViewPO.Show()
            End If

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                ' con.close()
            End If
        End Try
    End Sub

 

    Private Sub cboLocation_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboLocation.KeyUp
        If e.KeyCode = 13 Then
            txtRemark.Focus()
        ElseIf e.KeyCode = Keys.F2 Then
            frmViewPO.Show()
        End If
    End Sub

    Private Sub txtFind_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Call Load_Gride_Item3()
    End Sub

 
 

    Function Search_Supplier() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim Value As Double

        Try
            Sql = "select * from M09Supplier where  M09Name='" & Trim(cboLocation.Text) & "' and M09Active='A' and M09Loc_Code='" & _Comcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                _Loccode = Trim(M01.Tables(0).Rows(0)("M09Code"))
                Search_Supplier = True
            End If
            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Private Sub cmdAdd_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        Call Save_Data()
    End Sub

    Function Save_Data()
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
        Dim A As String
        Dim A1 As String
        Dim B As New ReportDocument

        Try
            If Search_Supplier() = True Then
            Else
                MsgBox("Please enter the correct Supplier Name", MsgBoxStyle.Information, "Information ........")
                cboLocation.ToggleDropdown()
                connection.Close()
                Exit Function
            End If


            If txtRemark.Text <> "" Then
            Else
                txtRemark.Text = " "
            End If

            nvcFieldList1 = "select * from T12PO_Header where T12PO_No='" & txtEntry.Text & "' and T12Loc_Code='" & _Comcode & "'"
            M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(M01) Then
                If Trim(M01.Tables(0).Rows(0)("T12Status")) = "A" Then
                    nvcFieldList1 = "UPDATE T12PO_Header SET T12Date='" & txtDate.Text & "',T12Supp_Code='" & _Loccode & "',T12Remark='" & txtRemark.Text & "' WHERE T12PO_No='" & txtEntry.Text & "' and T12Loc_Code='" & _Comcode & "'"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                    nvcFieldList1 = "DELETE FROM T13PO_Fluter WHERE T13Ref_No=" & M01.Tables(0).Rows(0)("T12Ref_No") & ""
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                    i = 0
                    For Each uRow As UltraGridRow In UltraGrid1.Rows
                        nvcFieldList1 = "Insert Into T13PO_Fluter(T13Ref_No,T13Item_Code,T13Qty,T13Rate,T13Count)" & _
                                                                " values('" & _EntryNo & "', '" & UltraGrid1.Rows(i).Cells(0).Text & "','" & UltraGrid1.Rows(i).Cells(3).Text & "','" & UltraGrid1.Rows(i).Cells(2).Text & "','" & i + 1 & "')"
                        ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                        i = i + 1
                    Next

                Else
                    MsgBox("You can not change this PO", MsgBoxStyle.Information, "Information ......")
                    connection.Close()
                    Exit Function
                End If
            Else
                Call Load_EntryNo()

                nvcFieldList1 = "select * from P01Parameter where P01Code='PO'"
                MB51 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                If isValidDataset(MB51) Then
                    _EntryNo = MB51.Tables(0).Rows(0)("P01LastNo")
                End If

                nvcFieldList1 = "UPDATE P01Parameter SET P01LastNo=P01LastNo +" & 1 & " WHERE P01Code='PO'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "Insert Into T12PO_Header(T12Ref_No,T12PO_No,T12Date,T12Supp_Code,T12Remark,T12Time,T12User,T12Status,T12Loc_Code)" & _
                                                             " values('" & _EntryNo & "', '" & (Trim(txtEntry.Text)) & "','" & txtDate.Text & "','" & _Loccode & "','" & txtRemark.Text & "','" & Now & "','" & strDisname & "','A','" & _Comcode & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                i = 0
                For Each uRow As UltraGridRow In UltraGrid1.Rows
                    nvcFieldList1 = "Insert Into T13PO_Fluter(T13Ref_No,T13Item_Code,T13Qty,T13Rate,T13Count)" & _
                                                            " values('" & _EntryNo & "', '" & UltraGrid1.Rows(i).Cells(0).Text & "','" & UltraGrid1.Rows(i).Cells(3).Text & "','" & UltraGrid1.Rows(i).Cells(2).Text & "','" & i + 1 & "')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                    i = i + 1
                Next
            End If
            transaction.Commit()
            'A = MsgBox("Are you sure you want to print this PO", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Print PO .....")
            'If A = vbYes Then
            '    A1 = ConfigurationManager.AppSettings("ReportPath") + "\PO.rpt"
            '    B.Load(A1.ToString)
            '   B.SetDatabaseLogon("sa", "tommya")
            '    'B.SetParameterValue("To", _To)
            '    'B.SetParameterValue("From", _From)
            '    '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
            '    frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
            '    frmReport.CrystalReportViewer1.DisplayToolbar = True
            '    frmReport.CrystalReportViewer1.SelectionFormula = "{T12PO_Header.T12PO_No}='" & txtEntry.Text & "' and {T12PO_Header.T12Status} <> 'Reject'"
            '    frmReport.Refresh()
            '    ' frmReport.CrystalReportViewer1.PrintReport()
            '    ' B.PrintToPrinter(1, True, 0, 0)
            '    frmReport.MdiParent = MDIMain
            '    frmReport.Show()
            'End If
            Call Clear_Text()
            Call Load_EntryNo()
            cboLocation.ToggleDropdown()
            connection.Close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.Close()
            End If
        End Try
    End Function


    Function Cancel_Records()
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
        Dim A As String
        Try
            A = MsgBox("Are you sure you want to cancel this PO", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Cancel PO .......")
            If A = vbYes Then
                nvcFieldList1 = "select * from T12PO_Header where T12PO_No='" & txtEntry.Text & "' and T12Status='A' and T12Loc_Code='" & _Comcode & "'"
                M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                If isValidDataset(M01) Then
                    nvcFieldList1 = "UPDATE T12PO_Header SET T12Status='Cancel' WHERE T12PO_No='" & txtEntry.Text & "' and T12Loc_Code='" & _Comcode & "'"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                    MsgBox("Transaction Canceled successfully", MsgBoxStyle.Information, "Information ......")
                    transaction.Commit()
                    Call Clear_Text()
                Else
                    MsgBox("You can not cancel this PO", MsgBoxStyle.Information, "Information .......")
                    connection.Close()
                    Exit Function
                End If
            End If

            connection.Close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.Close()
            End If
        End Try
    End Function
    Private Sub txtRate_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtRate.KeyUp
        Dim Value As Double
        If e.KeyCode = 13 Then
            If IsNumeric(txtRate.Text) Then
                Value = txtRate.Text
                txtRate.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtRate.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            End If
            Call Calculation()
            txtQty.Focus()
        ElseIf e.KeyCode = Keys.F2 Then
            frmViewPO.Show()
        End If
    End Sub

    Function Calculation()
        On Error Resume Next
        Dim Value As Double

        If IsNumeric(txtQty.Text) And IsNumeric(txtRate.Text) Then
            Value = CDbl(txtQty.Text) * CDbl(txtRate.Text)
            txtTotal.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            txtTotal.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
        End If
    End Function

    Private Sub txtRate_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRate.TextChanged
        Call Calculation()
    End Sub

    Private Sub txtQty_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtQty.TextChanged
        Call Calculation()
    End Sub

    Private Sub txtFind_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtFind.KeyUp
        If e.KeyCode = Keys.Escape Then
            OPR5.Visible = False
            cboCode.Focus()
        End If
    End Sub

    Private Sub txtFind_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFind.TextChanged
        Call Load_Gride_Item3()
    End Sub

    Private Sub UltraGrid3_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles UltraGrid3.KeyUp
        If e.KeyCode = 13 Then
            Dim _RowIndex As Integer

            _RowIndex = UltraGrid3.ActiveRow.Index
            cboCode.Text = Trim(UltraGrid3.Rows(_RowIndex).Cells(0).Text)
            Call Search_ItemName()
            txtFind.Text = ""
            OPR5.Visible = False
            txtRate.Focus()
        End If
    End Sub

    Private Sub UltraGrid3_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles UltraGrid3.MouseDoubleClick
        On Error Resume Next
        Dim _RowIndex As Integer

        _RowIndex = UltraGrid3.ActiveRow.Index
        cboCode.Text = Trim(UltraGrid3.Rows(_RowIndex).Cells(0).Text)
        Call Search_ItemName()
        txtFind.Text = ""
        OPR5.Visible = False
        txtRate.Focus()
    End Sub

    Private Sub UltraGrid1_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles UltraGrid1.AfterRowsDeleted
        Try
            Dim I As Integer
            Dim Value As Double
            I = 0
            For Each uRow As UltraGridRow In UltraGrid1.Rows
                Value = Value + UltraGrid1.Rows(I).Cells(4).Text
                I = I + 1
            Next

            txtNett.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            txtNett.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))


        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                'connection.Close()
            End If
        End Try
    End Sub

    Private Sub txtRemark_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtRemark.KeyUp
        If e.KeyCode = Keys.F2 Then
            frmViewPO.Show()
        End If
    End Sub

    Private Sub txtTotal_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTotal.KeyUp
        If e.KeyCode = Keys.F2 Then
            frmViewPO.Show()
        End If
    End Sub

    Private Sub cmdDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdDelete.Click
        Call Cancel_Records()
    End Sub

    Private Sub OPR0_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OPR0.Click

    End Sub

    Private Sub UltraButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton1.Click
        Dim A As String
        Dim A1 As String
        Dim B As New ReportDocument
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            A = MsgBox("Are you sure you want to print this PO", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Print PO .....")
            If A = vbYes Then
                Sql = "select * from T12PO_Header where T12PO_No='" & txtEntry.Text & "' and T12Loc_Code='" & _Comcode & "'"
                M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(M01) Then
                    If Trim(M01.Tables(0).Rows(0)("T12Status")) = "Approved" Then
                        A1 = ConfigurationManager.AppSettings("ReportPath") + "\PO.rpt"
                        B.Load(A1.ToString)
                        B.SetDatabaseLogon("sa", "tommya")
                        'B.SetParameterValue("To", _To)
                        'B.SetParameterValue("From", _From)
                        '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                        frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                        frmReport.CrystalReportViewer1.DisplayToolbar = True
                        frmReport.CrystalReportViewer1.SelectionFormula = "{T12PO_Header.T12PO_No}='" & txtEntry.Text & "' and {T12PO_Header.T12Status} <> 'Reject' and {T12PO_Header.T12Loc_Code}='" & _Comcode & "'"
                        frmReport.Refresh()
                        ' frmReport.CrystalReportViewer1.PrintReport()
                        ' B.PrintToPrinter(1, True, 0, 0)
                        frmReport.MdiParent = MDIMain
                        frmReport.Show()
                    Else
                        MsgBox("Please approved the PO", MsgBoxStyle.Information, "Information .....")
                    End If
                End If
            End If
            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                'connection.Close()
            End If
        End Try
    End Sub
End Class