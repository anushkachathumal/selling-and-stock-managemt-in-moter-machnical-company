Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports CrystalDecisions.CrystalReports.Engine
Public Class frmPO_Confirmation
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim c_dataCustomer1 As DataTable
    Dim _Category As String
    Dim _Comcode As String
    Dim _Loccode As String
    Dim _FromLocCode As String
    Dim _EntryNo As Integer


    Private Sub frmPO_Confirmation_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        _Comcode = ConfigurationManager.AppSettings("LOCCode")
        txtNett.ReadOnly = True
        txtNett.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
        txtDate.Text = Today
        Call Load_Gride2()
        txtSupplier.ReadOnly = True
        Call Load_PO()

    End Sub

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


    Function Load_PO()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select T12PO_No as [##] from T12PO_Header where T12Status='A' and T12Loc_Code='" & _Comcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboPO
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 110
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


    Function Clear_Text()
        Me.cboPO.Text = ""
        Me.txtRemark.Text = ""
        Me.txtNett.Text = ""
        Me.txtSupplier.Text = ""
        Call Load_Gride2()
        Call Load_PO()
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
            Sql = "select * from View_PO where T12PO_no='" & cboPO.Text & "' and Status<>'Canceled by User' and T12Loc_Code='" & _Comcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                ' txtRemark.Text = Trim(M01.Tables(0).Rows(0)("Remark"))
                txtDate.Text = Trim(M01.Tables(0).Rows(0)("date"))
                txtSupplier.Text = Trim(M01.Tables(0).Rows(0)("Supplier"))
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

    Private Sub cboPO_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboPO.AfterCloseUp
        Call Search_Records()
    End Sub

    Private Sub cboPO_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboPO.KeyUp
        If e.KeyCode = 13 Then
            Call Search_Records()
            txtRemark.Focus()
        End If
    End Sub

    Function Save_Approve()
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
        Dim A As String
        Try
            If txtRemark.Text <> "" Then
            Else
                txtRemark.Text = " "
            End If
            A = MsgBox("Are you sure you want to approved this PO", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Approved PO ....")
            If A = vbYes Then
                nvcFieldList1 = "select * from T12PO_Header where T12PO_No='" & Trim(cboPO.Text) & "' and T12Status<>'Cancel' and T12Loc_Code='" & _Comcode & "'"
                MB51 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                If isValidDataset(MB51) Then

                    nvcFieldList1 = "UPDATE T12PO_Header SET T12App_Remark='" & txtRemark.Text & "',T12App_Date='" & Today & "',T12Approved_By='" & strDisname & "',T12Status='Approved' WHERE  T12PO_No='" & cboPO.Text & "' and T12Loc_Code='" & _Comcode & "'"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                    MsgBox("Transaction completed", MsgBoxStyle.Information, "Information .......")
                    transaction.Commit()
                    Call Clear_Text()
                    Call Load_PO()
                    cboPO.ToggleDropdown()
                Else
                    MsgBox("This PO alrady canceled by user", MsgBoxStyle.Information, "Information ........")
                    Call Clear_Text()
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
    Private Sub cmdReset_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        Call Clear_Text()
    End Sub

    Private Sub cmdAdd_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        Call Save_Approve()
    End Sub


    Function Reject_Approve()
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
        Dim A As String
        Try
            If txtRemark.Text <> "" Then
            Else
                txtRemark.Text = " "
            End If
            A = MsgBox("Are you sure you want to reject this PO", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Approved PO ....")
            If A = vbYes Then
                nvcFieldList1 = "select * from T12PO_Header where T12PO_No='" & Trim(cboPO.Text) & "' and T12Status<>'Cancel' and T12Loc_Code='" & _Comcode & "'"
                MB51 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                If isValidDataset(MB51) Then

                    nvcFieldList1 = "UPDATE T12PO_Header SET T12App_Remark='" & txtRemark.Text & "',T12App_Date='" & Today & "',T12Approved_By='" & strDisname & "',T12Status='Reject' WHERE  T12PO_No='" & cboPO.Text & "' and T12Loc_Code='" & _Comcode & "'"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                    MsgBox("Transaction completed", MsgBoxStyle.Information, "Information .......")
                    transaction.Commit()
                    Call Clear_Text()
                    Call Load_PO()
                    cboPO.ToggleDropdown()
                Else
                    MsgBox("This PO alrady canceled by user", MsgBoxStyle.Information, "Information ........")
                    Call Clear_Text()
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

    Private Sub cmdDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdDelete.Click
        Call Reject_Approve()
    End Sub

    Private Sub UltraButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton1.Click
        Dim A As String
        Dim A1 As String
        Dim B As New ReportDocument

        Try
            A = MsgBox("Are you sure you want to print this PO", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Print PO .....")
            If A = vbYes Then
                A1 = ConfigurationManager.AppSettings("ReportPath") + "\PO.rpt"
                B.Load(A1.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                'B.SetParameterValue("To", _To)
                'B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{T12PO_Header.T12PO_No}='" & cboPO.Text & "' and {T12PO_Header.T12Status} <> 'Reject'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            End If

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                'connection.Close()
            End If
        End Try
    End Sub
End Class